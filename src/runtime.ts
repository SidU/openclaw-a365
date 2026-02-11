import type { PluginRuntime } from "openclaw/plugin-sdk";

/**
 * Module-level singleton for the A365 plugin runtime.
 *
 * This pattern is used because the OpenClaw plugin system initializes plugins
 * once and the runtime needs to be accessible from various modules (token.ts,
 * graph-tools.ts, etc.) without passing it through every function call.
 *
 * Trade-offs:
 * - Simple to use throughout the codebase
 * - Makes unit testing harder (use resetA365Runtime in tests)
 * - Not suitable for multiple plugin instances (not a current requirement)
 *
 * Alternative approaches for future consideration:
 * - Dependency injection container
 * - AsyncLocalStorage for request-scoped runtime
 * - Factory pattern with explicit runtime parameter
 */
let runtime: PluginRuntime | null = null;

/**
 * Set the A365 plugin runtime. Called once during plugin registration.
 */
export function setA365Runtime(next: PluginRuntime): void {
  runtime = next;
}

/**
 * Get the A365 plugin runtime.
 * @throws Error if runtime has not been initialized via setA365Runtime
 */
export function getA365Runtime(): PluginRuntime {
  if (!runtime) {
    throw new Error("A365 runtime not initialized - ensure plugin is registered before using runtime");
  }
  return runtime;
}

/**
 * Reset the runtime singleton. For testing purposes only.
 * @internal
 */
export function resetA365Runtime(): void {
  runtime = null;
}
