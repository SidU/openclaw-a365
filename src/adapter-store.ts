/**
 * Stores the CloudAdapter and Blueprint Client App ID for proactive messaging.
 *
 * The adapter is created during monitor startup and needed by outbound.ts
 * to call continueConversation() for proactive sends.
 *
 * The Blueprint Client App ID is the key insight from the SDK author:
 * continueConversation must be called with the Blueprint ID (not the bot's app ID)
 * so the SDK can correctly perform the T1/T2/AU token flow.
 */
import type { CloudAdapter } from "@microsoft/agents-hosting";

let storedAdapter: CloudAdapter | null = null;
let storedBlueprintClientId: string | null = null;

export function setAdapter(adapter: CloudAdapter): void {
  storedAdapter = adapter;
}

export function getAdapter(): CloudAdapter | null {
  return storedAdapter;
}

export function setBlueprintClientId(clientId: string): void {
  storedBlueprintClientId = clientId;
}

export function getBlueprintClientId(): string | null {
  return storedBlueprintClientId;
}
