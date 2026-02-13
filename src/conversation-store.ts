/**
 * Persistent conversation reference store for proactive messaging.
 *
 * Stores SDK ConversationReference objects (from activity.getConversationReference())
 * to a JSON file so they survive restarts. These references contain the agentic
 * identity metadata the SDK needs to acquire AU tokens for proactive sends.
 */
import { readFile, writeFile, mkdir } from "node:fs/promises";
import { join } from "node:path";
import { homedir } from "node:os";

/**
 * SDK ConversationReference shape (from @microsoft/agents-activity).
 * We use a loose type to avoid tight coupling to the SDK version.
 */
export type StoredConversationReference = {
  activityId?: string;
  user?: Record<string, unknown>;
  agent?: Record<string, unknown>;
  conversation: { id: string; [key: string]: unknown };
  channelId: string;
  serviceUrl?: string;
  locale?: string;
};

type ConversationStore = {
  references: Record<string, { ref: StoredConversationReference; updatedAt: number; userAadId?: string }>;
};

const STORE_DIR = join(homedir(), ".openclaw");
const STORE_PATH = join(STORE_DIR, "a365-conversations.json");

let storeCache: ConversationStore | null = null;

async function loadStore(): Promise<ConversationStore> {
  if (storeCache) return storeCache;
  try {
    const data = JSON.parse(await readFile(STORE_PATH, "utf-8")) as ConversationStore;
    storeCache = data;
    return data;
  } catch {
    storeCache = { references: {} };
    return storeCache;
  }
}

async function saveStore(store: ConversationStore): Promise<void> {
  storeCache = store;
  await mkdir(STORE_DIR, { recursive: true });
  await writeFile(STORE_PATH, JSON.stringify(store, null, 2), "utf-8");
}

/**
 * Save a conversation reference for proactive messaging.
 */
export async function saveConversationReference(
  ref: StoredConversationReference,
  userAadId?: string,
): Promise<void> {
  const store = await loadStore();
  const key = ref.conversation.id;
  store.references[key] = { ref, updatedAt: Date.now(), userAadId };
  await saveStore(store);
}

/**
 * Get a stored conversation reference by conversation ID.
 */
export async function getConversationReference(
  conversationId: string,
): Promise<StoredConversationReference | undefined> {
  const store = await loadStore();
  return store.references[conversationId]?.ref;
}

/**
 * Get a stored conversation reference by user AAD ID.
 */
export async function getConversationReferenceByUser(
  userAadId: string,
): Promise<StoredConversationReference | undefined> {
  const store = await loadStore();
  let best: { ref: StoredConversationReference; updatedAt: number } | undefined;
  for (const entry of Object.values(store.references)) {
    if (entry.userAadId === userAadId && (!best || entry.updatedAt > best.updatedAt)) {
      best = entry;
    }
  }
  return best?.ref;
}

/**
 * List all stored conversation references.
 */
export async function listConversationReferences(): Promise<
  Array<{ conversationId: string; updatedAt: number; userAadId?: string }>
> {
  const store = await loadStore();
  return Object.entries(store.references).map(([id, entry]) => ({
    conversationId: id,
    updatedAt: entry.updatedAt,
    userAadId: entry.userAadId,
  }));
}

/**
 * Delete a stored conversation reference.
 */
export async function deleteConversationReference(conversationId: string): Promise<void> {
  const store = await loadStore();
  delete store.references[conversationId];
  await saveStore(store);
}

/**
 * Clear all stored conversation references.
 */
export async function clearConversationReferences(): Promise<void> {
  await saveStore({ references: {} });
}
