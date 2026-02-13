# OpenClaw A365 Channel Docker Image
# Provides native Microsoft 365 Agents integration with Graph API tools

FROM node:22-alpine

LABEL org.opencontainers.image.title="OpenClaw A365 Channel"
LABEL org.opencontainers.image.description="Microsoft 365 Agents channel with native Graph API tools"
LABEL org.opencontainers.image.source="https://github.com/openclaw/openclaw"

# Install dependencies for native modules and network policy enforcement
RUN apk add --no-cache python3 make g++ git cmake iptables

WORKDIR /app

# Copy package files
COPY package.json pnpm-lock.yaml* ./

# Install pnpm and dependencies
# Skip node-llama-cpp postinstall as we use Anthropic API, not local LLMs
RUN npm install -g pnpm@9 && \
    NODE_LLAMA_CPP_SKIP_DOWNLOAD=true pnpm install --ignore-scripts

# Copy application code
COPY . .

# OpenClaw handles TypeScript compilation via tsx at runtime

# Install plugin into default state dir during build (can't use /app/data since plugin source is /app)
RUN mkdir -p /root/.openclaw && \
    echo '{"gateway":{"mode":"local"}}' > /root/.openclaw/openclaw.json

# Install the A365 plugin into OpenClaw and enable it
RUN pnpm openclaw plugins install . && \
    pnpm openclaw plugins enable a365

# Now copy the full config with A365 channel enabled
COPY config/openclaw.json /root/.openclaw/openclaw.json

# At runtime, OPENCLAW_STATE_DIR=/app/data (set via docker-compose)
# The entrypoint seeds /app/data from /root/.openclaw on first start

# Copy network policy script
COPY scripts/network-policy.sh /app/scripts/network-policy.sh
RUN chmod +x /app/scripts/network-policy.sh

# Create entrypoint script with state seeding, network policy, and model configuration
RUN printf '%s\n' '#!/bin/sh' > /app/entrypoint.sh && \
    printf '%s\n' '' >> /app/entrypoint.sh && \
    printf '%s\n' '# Seed persistent state dir from build-time install on first start' >> /app/entrypoint.sh && \
    printf '%s\n' 'if [ -n "$OPENCLAW_STATE_DIR" ] && [ ! -f "$OPENCLAW_STATE_DIR/openclaw.json" ]; then' >> /app/entrypoint.sh && \
    printf '%s\n' '  echo "=== First start: seeding $OPENCLAW_STATE_DIR from build ==="' >> /app/entrypoint.sh && \
    printf '%s\n' '  cp -a /root/.openclaw/. "$OPENCLAW_STATE_DIR/"' >> /app/entrypoint.sh && \
    printf '%s\n' 'fi' >> /app/entrypoint.sh && \
    printf '%s\n' '' >> /app/entrypoint.sh && \
    printf '%s\n' '# Apply network policy (if not unrestricted)' >> /app/entrypoint.sh && \
    printf '%s\n' '/app/scripts/network-policy.sh' >> /app/entrypoint.sh && \
    printf '%s\n' '' >> /app/entrypoint.sh && \
    printf '%s\n' 'echo "=== Running doctor --fix ==="' >> /app/entrypoint.sh && \
    printf '%s\n' 'pnpm openclaw doctor --fix || echo "Doctor completed with warnings"' >> /app/entrypoint.sh && \
    printf '%s\n' '' >> /app/entrypoint.sh && \
    printf '%s\n' 'echo "=== Configuring models ==="' >> /app/entrypoint.sh && \
    printf '%s\n' 'if [ -n "$OPENCLAW_MODEL" ]; then' >> /app/entrypoint.sh && \
    printf '%s\n' '  echo "Setting primary model: $OPENCLAW_MODEL"' >> /app/entrypoint.sh && \
    printf '%s\n' '  pnpm openclaw models set "$OPENCLAW_MODEL" || echo "Warning: Could not set model"' >> /app/entrypoint.sh && \
    printf '%s\n' 'fi' >> /app/entrypoint.sh && \
    printf '%s\n' '' >> /app/entrypoint.sh && \
    printf '%s\n' 'if [ -n "$OPENCLAW_FALLBACK_MODELS" ]; then' >> /app/entrypoint.sh && \
    printf '%s\n' '  echo "Configuring fallback models: $OPENCLAW_FALLBACK_MODELS"' >> /app/entrypoint.sh && \
    printf '%s\n' '  pnpm openclaw models fallbacks clear 2>/dev/null || true' >> /app/entrypoint.sh && \
    printf '%s\n' '  echo "$OPENCLAW_FALLBACK_MODELS" | tr "," "\n" | while read model; do' >> /app/entrypoint.sh && \
    printf '%s\n' '    model=$(echo "$model" | xargs)' >> /app/entrypoint.sh && \
    printf '%s\n' '    if [ -n "$model" ]; then' >> /app/entrypoint.sh && \
    printf '%s\n' '      echo "  Adding fallback: $model"' >> /app/entrypoint.sh && \
    printf '%s\n' '      pnpm openclaw models fallbacks add "$model" || echo "  Warning: Could not add fallback"' >> /app/entrypoint.sh && \
    printf '%s\n' '    fi' >> /app/entrypoint.sh && \
    printf '%s\n' '  done' >> /app/entrypoint.sh && \
    printf '%s\n' 'fi' >> /app/entrypoint.sh && \
    printf '%s\n' '' >> /app/entrypoint.sh && \
    printf '%s\n' 'echo "=== Model configuration ==="' >> /app/entrypoint.sh && \
    printf '%s\n' 'pnpm openclaw models status 2>/dev/null || true' >> /app/entrypoint.sh && \
    printf '%s\n' '' >> /app/entrypoint.sh && \
    printf '%s\n' 'echo "=== Starting gateway ==="' >> /app/entrypoint.sh && \
    printf '%s\n' 'exec pnpm openclaw gateway' >> /app/entrypoint.sh && \
    chmod +x /app/entrypoint.sh

# Expose A365 messaging port
EXPOSE 3978

# Health check
HEALTHCHECK --interval=30s --timeout=10s --start-period=10s --retries=3 \
  CMD wget --no-verbose --tries=1 --spider http://localhost:3978/health || exit 1

# Default command - run entrypoint that does doctor --fix then gateway
ENTRYPOINT ["/bin/sh", "/app/entrypoint.sh"]
