import { readFile, writeFile, mkdir } from "node:fs/promises";
import { existsSync } from "node:fs";
import { join } from "node:path";
import { homedir } from "node:os";
import type { StoredConversationReference } from "./types.js";
import { getA365Runtime } from "./runtime.js";

/**
 * File-based conversation reference store for proactive messaging.
 *
 * Stores conversation references in a JSON file so that the bot can send
 * proactive messages later (e.g., from cron jobs, async task completion).
 *
 * Storage location: ~/.openclaw/a365-conversations.json
 *
 * TODO: For multi-instance deployments, consider using a shared database
 * (Redis, SQLite, etc.) instead of file-based storage.
 */

type ConversationStore = {
  version: number;
  conversations: Record<string, StoredConversationReference>;
};

const STORE_VERSION = 1;
const STORE_DIR = join(homedir(), ".openclaw");
const STORE_PATH = join(STORE_DIR, "a365-conversations.json");

// In-memory cache to avoid reading file on every lookup
let memoryCache: ConversationStore | null = null;
let cacheLoadedAt = 0;
const CACHE_TTL_MS = 30_000; // Reload from disk every 30 seconds

/**
 * Ensure the storage directory exists.
 */
async function ensureStoreDir(): Promise<void> {
  if (!existsSync(STORE_DIR)) {
    await mkdir(STORE_DIR, { recursive: true });
  }
}

/**
 * Load the conversation store from disk.
 */
async function loadStore(): Promise<ConversationStore> {
  // Return cached version if fresh
  if (memoryCache && Date.now() - cacheLoadedAt < CACHE_TTL_MS) {
    return memoryCache;
  }

  try {
    if (existsSync(STORE_PATH)) {
      const data = await readFile(STORE_PATH, "utf-8");
      const parsed = JSON.parse(data) as ConversationStore;

      // Validate version
      if (parsed.version === STORE_VERSION && parsed.conversations) {
        memoryCache = parsed;
        cacheLoadedAt = Date.now();
        return parsed;
      }
    }
  } catch (err) {
    const log = getA365Runtime().logging.getChildLogger({ name: "conversation-store" });
    log.warn("Failed to load conversation store, starting fresh", { error: String(err) });
  }

  // Return empty store
  const empty: ConversationStore = { version: STORE_VERSION, conversations: {} };
  memoryCache = empty;
  cacheLoadedAt = Date.now();
  return empty;
}

/**
 * Save the conversation store to disk.
 */
async function saveStore(store: ConversationStore): Promise<void> {
  await ensureStoreDir();
  await writeFile(STORE_PATH, JSON.stringify(store, null, 2), "utf-8");
  memoryCache = store;
  cacheLoadedAt = Date.now();
}

/**
 * Save a conversation reference for later proactive messaging.
 *
 * Call this when receiving a message to store the context needed
 * to send messages back to this conversation later.
 */
export async function saveConversationReference(
  ref: StoredConversationReference,
): Promise<void> {
  const log = getA365Runtime().logging.getChildLogger({ name: "conversation-store" });

  try {
    const store = await loadStore();
    store.conversations[ref.conversationId] = ref;
    await saveStore(store);
    log.debug("Saved conversation reference", { conversationId: ref.conversationId });
  } catch (err) {
    log.error("Failed to save conversation reference", {
      conversationId: ref.conversationId,
      error: String(err),
    });
    throw err;
  }
}

/**
 * Get a stored conversation reference by conversation ID.
 *
 * Returns null if no reference is found for this conversation.
 */
export async function getConversationReference(
  conversationId: string,
): Promise<StoredConversationReference | null> {
  const log = getA365Runtime().logging.getChildLogger({ name: "conversation-store" });

  try {
    const store = await loadStore();
    const storeKeys = Object.keys(store.conversations);
    log.debug(`Looking up conversation: lookupKey=${conversationId} storeKeyCount=${storeKeys.length} storeKeys=${JSON.stringify(storeKeys.slice(0, 3))}`);

    const ref = store.conversations[conversationId];

    if (ref) {
      log.debug(`Found conversation reference: ${conversationId}`);
      return ref;
    }

    log.debug(`No conversation reference found: lookupKey=${conversationId} availableKeys=${JSON.stringify(storeKeys)}`);
    return null;
  } catch (err) {
    log.error("Failed to get conversation reference", {
      conversationId,
      error: String(err),
    });
    return null;
  }
}

/**
 * Get a conversation reference by user AAD ID.
 *
 * Useful when you need to message a user but only have their AAD ID.
 * Returns the most recently updated conversation with this user.
 */
export async function getConversationReferenceByUser(
  userAadId: string,
): Promise<StoredConversationReference | null> {
  const log = getA365Runtime().logging.getChildLogger({ name: "conversation-store" });

  try {
    const store = await loadStore();
    let best: StoredConversationReference | null = null;

    for (const ref of Object.values(store.conversations)) {
      if (ref.userAadId === userAadId) {
        // Prefer DM conversations and more recent ones
        if (!best || (!ref.isGroup && best.isGroup) || ref.updatedAt > best.updatedAt) {
          best = ref;
        }
      }
    }

    if (best) {
      log.debug("Found conversation reference by user", { userAadId, conversationId: best.conversationId });
    } else {
      log.debug("No conversation reference found for user", { userAadId });
    }

    return best;
  } catch (err) {
    log.error("Failed to get conversation reference by user", {
      userAadId,
      error: String(err),
    });
    return null;
  }
}

/**
 * Delete a conversation reference.
 *
 * Call this if you know a conversation is no longer valid
 * (e.g., bot was removed from the conversation).
 */
export async function deleteConversationReference(conversationId: string): Promise<void> {
  const log = getA365Runtime().logging.getChildLogger({ name: "conversation-store" });

  try {
    const store = await loadStore();
    if (store.conversations[conversationId]) {
      delete store.conversations[conversationId];
      await saveStore(store);
      log.debug("Deleted conversation reference", { conversationId });
    }
  } catch (err) {
    log.error("Failed to delete conversation reference", {
      conversationId,
      error: String(err),
    });
  }
}

/**
 * List all stored conversation references.
 *
 * Useful for debugging or admin purposes.
 */
export async function listConversationReferences(): Promise<StoredConversationReference[]> {
  const store = await loadStore();
  return Object.values(store.conversations);
}

/**
 * Clear all stored conversation references.
 *
 * Use with caution - this will prevent proactive messaging to all conversations
 * until new messages are received.
 */
export async function clearConversationReferences(): Promise<void> {
  const log = getA365Runtime().logging.getChildLogger({ name: "conversation-store" });
  const empty: ConversationStore = { version: STORE_VERSION, conversations: {} };
  await saveStore(empty);
  log.info("Cleared all conversation references");
}
