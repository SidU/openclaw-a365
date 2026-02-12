import type { OpenClawPluginApi } from "openclaw/plugin-sdk";
import { emptyPluginConfigSchema } from "openclaw/plugin-sdk";
import { a365Plugin } from "./src/channel.js";
import { setA365Runtime } from "./src/runtime.js";

// Re-export monitor for external use
export { monitorA365Provider, buildConversationReference } from "./src/monitor.js";
export { createGraphTools } from "./src/graph-tools.js";
export { sendMessageA365, sendAdaptiveCardA365 } from "./src/outbound.js";
export {
  saveConversationReference,
  getConversationReference,
  getConversationReferenceByUser,
  deleteConversationReference,
  listConversationReferences,
  clearConversationReferences,
} from "./src/conversation-store.js";
export type { A365Config, A365MessageMetadata, GraphCalendarEvent, StoredConversationReference } from "./src/types.js";

const plugin = {
  id: "a365",
  name: "Microsoft 365 Agents",
  description: "A365 channel plugin with native Graph API tools for calendar and email",
  configSchema: emptyPluginConfigSchema(),
  register(api: OpenClawPluginApi) {
    setA365Runtime(api.runtime);
    api.registerChannel({ plugin: a365Plugin });
  },
};

export default plugin;
