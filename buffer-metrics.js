#!/usr/bin/env node
/**
 * Pulls current-vs-prior-period post metrics for all connected
 * Instagram/TikTok/Facebook channels and prints clean JSON grouped by
 * channel, ready for generate_report.py's month-over-month comparison
 * slides.
 *
 * Requires Node 18+ (built-in fetch). No external dependencies.
 *
 * NETWORK: the environment this runs in must allow outbound access to
 * api.buffer.com (see SETUP.md, step 3) -- a restrictive/default network
 * policy blocks this silently and the failure looks exactly like a bad API
 * key ("403 Forbidden") rather than a network error, unless you read the
 * response body this script now surfaces on failure.
 *
 * Env vars:
 *   SALONVERITAS_BUFFER_API_KEY  (required) Buffer API access token
 *   BUFFER_GRAPHQL_URL         (optional) defaults to https://api.buffer.com/graphql
 *   BUFFER_ORGANIZATION_ID     (optional) only needed if the account belongs
 *                              to more than one organization -- otherwise the
 *                              account's first organization is used.
 *
 * Period design (locked in after live diagnostics on 2026-07-18):
 *   - Current period is always the last 30 days.
 *   - Prior period width is PER CHANNEL, not uniform:
 *       - Instagram: 60 days (day 30-90 back). A straight 30-day-wide prior
 *         window only had 1 post -- too thin a sample to compare against.
 *         Widening to 60 days gave 10 posts, comparable to the current
 *         period's 7.
 *       - TikTok / Facebook: 30 days (day 30-60 back) -- posting volume was
 *         high enough there for a same-width comparison to be fair.
 *   Both periods are derived identically (summed from per-post `metrics`),
 *   never from `aggregatedPostMetrics`, so current and prior are always
 *   computed the same way -- see the retention note below for why.
 *
 * Schema verified by introspection/live testing against a real token
 * (2026-07-17 and 2026-07-18):
 *   - `aggregatedPostMetrics` (a channel-level KPI summary field) is capped
 *     by Buffer's Free plan to the last ~30 days -- querying further back
 *     errors with "Free-plan Insights are limited to the last 31 days of
 *     history." Since every prior-period window here starts more than 30
 *     days back, this field can never be used for the prior period, so it
 *     isn't used for the current period either (keeps both periods on one
 *     consistent code path). It has been dropped from this script entirely.
 *   - The `posts` query (per-post breakdown, each with its own `metrics`
 *     list) is NOT subject to that cap -- real per-post metrics were
 *     confirmed back to late March 2026 (~114 days) for all three channels,
 *     easily covering both periods needed here.
 *   - Metric field names are NOT uniform across platforms: Facebook reports
 *     "impressions", Instagram/TikTok report "views". Instagram posts also
 *     carry "saves"/"follows"; Facebook and TikTok posts do not.
 *   - There is no follower-count or follower-history field anywhere in this
 *     API (confirmed via full schema introspection of Channel and every
 *     per-service ChannelMetadata type) -- a true follower count/growth
 *     series cannot be built from this data on any platform.
 */

const API_KEY = process.env.SALONVERITAS_BUFFER_API_KEY;
if (!API_KEY) {
  console.error('Missing required env var: SALONVERITAS_BUFFER_API_KEY');
  process.exit(1);
}

const GRAPHQL_URL = process.env.BUFFER_GRAPHQL_URL || 'https://api.buffer.com/graphql';
const TARGET_SERVICES = ['instagram', 'tiktok', 'facebook'];
const CURRENT_PERIOD_DAYS = 30;
const PRIOR_PERIOD_DAYS = { instagram: 60, tiktok: 30, facebook: 30 };
const POSTS_PAGE_SIZE = 50;
const POSTS_MAX_PAGES = 20; // safety cap against a runaway pagination loop

async function graphqlRequest(query, variables) {
  const res = await fetch(GRAPHQL_URL, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      Authorization: `Bearer ${API_KEY}`,
    },
    body: JSON.stringify({ query, variables }),
  });

  if (!res.ok) {
    const bodyText = await res.text().catch(() => '');
    throw new Error(`GraphQL request failed: ${res.status} ${res.statusText} -- ${bodyText}`);
  }

  const json = await res.json();
  if (json.errors && json.errors.length) {
    throw new Error(json.errors.map((e) => e.message).join('; '));
  }

  return json.data;
}

function isoDateTime(date) {
  return date.toISOString();
}

function metricsListToDict(metrics) {
  const dict = {};
  for (const m of metrics || []) dict[m.type] = m.value;
  return dict;
}

async function fetchOrganizationId() {
  if (process.env.BUFFER_ORGANIZATION_ID) return process.env.BUFFER_ORGANIZATION_ID;
  const data = await graphqlRequest('{ account { organizations { id } } }', {});
  const org = data.account.organizations[0];
  if (!org) throw new Error('Buffer account has no organizations');
  return org.id;
}

async function fetchConnectedChannels(organizationId) {
  const data = await graphqlRequest(
    `query Channels($input: ChannelsInput!) {
      channels(input: $input) { id name displayName service }
    }`,
    { input: { organizationId } }
  );
  return data.channels.filter((c) => TARGET_SERVICES.includes(c.service));
}

async function fetchPosts(organizationId, channelId, startDateTime, endDateTime) {
  const posts = [];
  let after = null;
  for (let page = 0; page < POSTS_MAX_PAGES; page++) {
    const data = await graphqlRequest(
      `query Posts($input: PostsInput!, $first: Int, $after: String) {
        posts(input: $input, first: $first, after: $after) {
          pageInfo { hasNextPage endCursor }
          edges { node { id text sentAt metrics { type name value unit } } }
        }
      }`,
      {
        input: {
          organizationId,
          filter: {
            channelIds: [channelId],
            status: ['sent'],
            dueAt: { start: startDateTime, end: endDateTime },
          },
          sort: [{ field: 'dueAt', direction: 'desc' }],
        },
        first: POSTS_PAGE_SIZE,
        after,
      }
    );

    for (const { node } of data.posts.edges) {
      // dueAt (the schedule filter) can drift slightly from sentAt (actual
      // publish time); re-check sentAt so the window stays accurate.
      if (node.sentAt && node.sentAt >= startDateTime && node.sentAt <= endDateTime) {
        posts.push({
          postId: node.id,
          text: node.text,
          sentAt: node.sentAt,
          metrics: metricsListToDict(node.metrics),
        });
      }
    }

    if (!data.posts.pageInfo.hasNextPage) break;
    after = data.posts.pageInfo.endCursor;
  }
  return posts;
}

async function main() {
  const now = new Date();
  const currentStart = new Date(now.getTime() - CURRENT_PERIOD_DAYS * 24 * 60 * 60 * 1000);
  const organizationId = await fetchOrganizationId();
  const channels = await fetchConnectedChannels(organizationId);

  const grouped = {};
  for (const service of TARGET_SERVICES) grouped[service] = [];

  await Promise.all(
    channels.map(async (channel) => {
      const priorDays = PRIOR_PERIOD_DAYS[channel.service] ?? CURRENT_PERIOD_DAYS;
      const priorStart = new Date(currentStart.getTime() - priorDays * 24 * 60 * 60 * 1000);
      const entry = { channelId: channel.id, name: channel.displayName || channel.name };

      try {
        // One paginated fetch spanning both periods, then split client-side
        // -- avoids double-fetching the boundary and keeps both periods on
        // an identical code path.
        const allPosts = await fetchPosts(
          organizationId, channel.id, isoDateTime(priorStart), isoDateTime(now)
        );
        entry.current = {
          range: { startDate: isoDateTime(currentStart), endDate: isoDateTime(now), days: CURRENT_PERIOD_DAYS },
          posts: allPosts.filter((p) => p.sentAt >= isoDateTime(currentStart)),
        };
        entry.prior = {
          range: { startDate: isoDateTime(priorStart), endDate: isoDateTime(currentStart), days: priorDays },
          posts: allPosts.filter(
            (p) => p.sentAt >= isoDateTime(priorStart) && p.sentAt < isoDateTime(currentStart)
          ),
        };
      } catch (err) {
        entry.error = err.message;
      }

      grouped[channel.service].push(entry);
    })
  );

  console.log(JSON.stringify({ channels: grouped }, null, 2));
}

main().catch((err) => {
  console.error(err.message);
  process.exit(1);
});
