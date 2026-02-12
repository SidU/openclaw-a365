/**
 * Store for the CloudAdapter instance.
 *
 * The outbound adapter uses this to construct a TurnContext for proactive
 * messaging, injecting a ConnectorClient authenticated with the agent's
 * own T1/T2/User FIC token.
 */

import type { CloudAdapter } from "@microsoft/agents-hosting";

let storedAdapter: CloudAdapter | null = null;

/**
 * Store the CloudAdapter for proactive messaging.
 */
export function setAdapter(adapter: CloudAdapter): void {
  storedAdapter = adapter;
}

/**
 * Get the stored CloudAdapter.
 */
export function getAdapter(): CloudAdapter | null {
  return storedAdapter;
}
