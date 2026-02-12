#!/usr/bin/env node
/**
 * Test script for proactive messaging.
 * Run inside the container: node /app/scripts/test-proactive.mjs
 */

const APP_ID = process.env.A365_APP_ID || process.env.MicrosoftAppId;
const APP_SECRET = process.env.A365_APP_PASSWORD || process.env.MicrosoftAppPassword;
const TENANT_ID = process.env.A365_TENANT_ID || process.env.MicrosoftAppTenantId;
const AA_INSTANCE_ID = process.env.AA_INSTANCE_ID;
const AGENT_IDENTITY = process.env.AGENT_IDENTITY;
const BLUEPRINT_CLIENT_APP_ID = process.env.BLUEPRINT_CLIENT_APP_ID || APP_ID;
const BLUEPRINT_CLIENT_SECRET = process.env.BLUEPRINT_CLIENT_SECRET || APP_SECRET;

console.log("=== Config ===");
console.log("APP_ID:", APP_ID);
console.log("TENANT_ID:", TENANT_ID);
console.log("AA_INSTANCE_ID:", AA_INSTANCE_ID);
console.log("AGENT_IDENTITY:", AGENT_IDENTITY);
console.log("BLUEPRINT_CLIENT_APP_ID:", BLUEPRINT_CLIENT_APP_ID);

if (!APP_ID || !APP_SECRET || !TENANT_ID || !AA_INSTANCE_ID || !AGENT_IDENTITY) {
  console.error("Missing required env vars");
  process.exit(1);
}

// Load conversation store
import { readFile } from "node:fs/promises";
import { join } from "node:path";
import { homedir } from "node:os";

const storePath = join(homedir(), ".openclaw", "a365-conversations.json");
let conversationId, serviceUrl;

try {
  const store = JSON.parse(await readFile(storePath, "utf-8"));
  const convs = Object.values(store.conversations);
  const latest = convs.sort((a, b) => b.updatedAt - a.updatedAt)[0];
  conversationId = latest.conversationId;
  serviceUrl = latest.serviceUrl;
  console.log("conversationId:", conversationId);
  console.log("serviceUrl:", serviceUrl);
} catch (err) {
  console.error("Failed to load conversation store:", err.message);
  process.exit(1);
}

// ============================================================
// Approach 1: Manual T1/T2/User FIC (same as token.ts) + ConnectorClient
// ============================================================
console.log("\n=== Approach 1: Manual T1/T2 + ConnectorClient ===");

const tokenEndpoint = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;

async function manualT1T2(scope) {
  // T1
  const t1Body = new URLSearchParams({
    scope: "api://AzureAdTokenExchange/.default",
    client_id: BLUEPRINT_CLIENT_APP_ID,
    grant_type: "client_credentials",
    client_secret: BLUEPRINT_CLIENT_SECRET,
    fmi_path: AA_INSTANCE_ID,
  });
  const t1Resp = await fetch(tokenEndpoint, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: t1Body.toString(),
  });
  if (!t1Resp.ok) throw new Error(`T1 failed: ${await t1Resp.text()}`);
  const t1 = await t1Resp.json();
  console.log("  T1 OK");

  // T2
  const t2Body = new URLSearchParams({
    scope: "api://AzureAdTokenExchange/.default",
    client_id: AA_INSTANCE_ID,
    grant_type: "client_credentials",
    client_assertion_type: "urn:ietf:params:oauth:client-assertion-type:jwt-bearer",
    client_assertion: t1.access_token,
  });
  const t2Resp = await fetch(tokenEndpoint, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: t2Body.toString(),
  });
  if (!t2Resp.ok) throw new Error(`T2 failed: ${await t2Resp.text()}`);
  const t2 = await t2Resp.json();
  console.log("  T2 OK");

  // User FIC
  const userBody = new URLSearchParams({
    scope,
    client_id: AA_INSTANCE_ID,
    grant_type: "user_fic",
    client_assertion_type: "urn:ietf:params:oauth:client-assertion-type:jwt-bearer",
    client_assertion: t1.access_token,
    username: AGENT_IDENTITY,
    user_federated_identity_credential: t2.access_token,
  });
  const userResp = await fetch(tokenEndpoint, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: userBody.toString(),
  });
  if (!userResp.ok) {
    const errText = await userResp.text();
    throw new Error(`User FIC failed (scope=${scope}): ${errText}`);
  }
  const user = await userResp.json();
  console.log(`  User FIC OK (scope=${scope})`);
  return user.access_token;
}

// Test with different scopes
const testScopes = [
  "https://graph.microsoft.com/.default",
  "5a807f24-c9de-44ee-a3a7-329e88a00ffc/.default",
  "https://api.botframework.com/.default",
];

const { ConnectorClient } = await import("@microsoft/agents-hosting");

for (const scope of testScopes) {
  console.log(`\nTrying manual T1/T2/User FIC with scope: ${scope}`);
  try {
    const token = await manualT1T2(scope);

    // Try sending with this token
    console.log("  Sending with ConnectorClient...");
    const client = ConnectorClient.createClientWithToken(serviceUrl, token);
    const result = await client.sendToConversation(conversationId, {
      type: "message",
      text: `Proactive test (scope: ${scope})`,
    });
    console.log("  SENT! Result:", JSON.stringify(result?.data || result));
    break; // Stop on first success
  } catch (err) {
    const detail = err.response?.data?.error_description || err.response?.data || err.message;
    console.error("  FAILED:", typeof detail === 'string' ? detail.substring(0, 200) : JSON.stringify(detail).substring(0, 200));
  }
}

// ============================================================
// Approach 2: Compare SDK user_id vs manual username
// ============================================================
console.log("\n=== Approach 2: SDK getAgenticUserToken comparison ===");
console.log("The SDK uses 'user_id' field, our manual flow uses 'username'.");
console.log("Testing if 'user_id' is the issue...");

// Manual User FIC with user_id instead of username (like SDK does)
try {
  // T1
  const t1Body = new URLSearchParams({
    scope: "api://AzureAdTokenExchange/.default",
    client_id: BLUEPRINT_CLIENT_APP_ID,
    grant_type: "client_credentials",
    client_secret: BLUEPRINT_CLIENT_SECRET,
    fmi_path: AA_INSTANCE_ID,
  });
  const t1Resp = await fetch(tokenEndpoint, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: t1Body.toString(),
  });
  const t1 = await t1Resp.json();

  // T2
  const t2Body = new URLSearchParams({
    scope: "api://AzureAdTokenExchange/.default",
    client_id: AA_INSTANCE_ID,
    grant_type: "client_credentials",
    client_assertion_type: "urn:ietf:params:oauth:client-assertion-type:jwt-bearer",
    client_assertion: t1.access_token,
  });
  const t2Resp = await fetch(tokenEndpoint, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: t2Body.toString(),
  });
  const t2 = await t2Resp.json();

  // User FIC with "user_id" instead of "username" (SDK style)
  const scope = "https://graph.microsoft.com/.default";
  const userBody = new URLSearchParams({
    scope,
    client_id: AA_INSTANCE_ID,
    grant_type: "user_fic",
    client_assertion_type: "urn:ietf:params:oauth:client-assertion-type:jwt-bearer",
    client_assertion: t1.access_token,
    user_id: AGENT_IDENTITY,  // SDK uses user_id, not username
    user_federated_identity_credential: t2.access_token,
  });
  const userResp = await fetch(tokenEndpoint, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: userBody.toString(),
  });
  if (!userResp.ok) {
    console.error("  user_id style FAILED:", (await userResp.text()).substring(0, 200));
  } else {
    console.log("  user_id style OK! (The SDK's field name is fine)");
  }
} catch (err) {
  console.error("  Error:", err.message);
}
