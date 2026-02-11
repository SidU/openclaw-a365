#!/bin/sh
# Network Policy Enforcement Script
# Applies iptables rules based on NETWORK_MODE environment variable
#
# Modes:
#   unrestricted - No restrictions (default)
#   restricted   - Only essential domains (Microsoft, configured LLM provider)
#   allowlist    - Essential + NETWORK_ALLOWLIST domains

set -e

NETWORK_MODE="${NETWORK_MODE:-unrestricted}"
NETWORK_ALLOWLIST="${NETWORK_ALLOWLIST:-}"

# Essential domains that are always needed
ESSENTIAL_DOMAINS="
login.microsoftonline.com
graph.microsoft.com
smba.trafficmanager.net
directline.botframework.com
api.botframework.com
"

# LLM provider domains based on configured keys
get_llm_domains() {
  domains=""

  # Anthropic
  if [ -n "$ANTHROPIC_API_KEY" ]; then
    domains="$domains api.anthropic.com"
  fi

  # OpenAI
  if [ -n "$OPENAI_API_KEY" ]; then
    domains="$domains api.openai.com"
  fi

  # OpenRouter
  if [ -n "$OPENROUTER_API_KEY" ]; then
    domains="$domains openrouter.ai"
  fi

  # Azure OpenAI (extract hostname from endpoint)
  if [ -n "$AZURE_OPENAI_ENDPOINT" ]; then
    azure_host=$(echo "$AZURE_OPENAI_ENDPOINT" | sed -E 's|https?://([^/]+).*|\1|')
    if [ -n "$azure_host" ]; then
      domains="$domains $azure_host"
    fi
  fi

  echo "$domains"
}

# Resolve domain to IP addresses (handles multiple IPs)
resolve_domain() {
  domain="$1"
  # Use getent for reliable resolution, fall back to nslookup
  if command -v getent >/dev/null 2>&1; then
    getent ahosts "$domain" 2>/dev/null | awk '{print $1}' | sort -u | grep -v ':'
  else
    nslookup "$domain" 2>/dev/null | awk '/^Address: / { print $2 }' | grep -v ':'
  fi
}

# Apply iptables rules for a domain
allow_domain() {
  domain="$1"
  echo "  Allowing: $domain"

  ips=$(resolve_domain "$domain")
  if [ -z "$ips" ]; then
    echo "    Warning: Could not resolve $domain"
    return
  fi

  for ip in $ips; do
    echo "    -> $ip"
    iptables -A OUTPUT -d "$ip" -j ACCEPT 2>/dev/null || true
  done
}

echo "=== Network Policy: $NETWORK_MODE ==="

if [ "$NETWORK_MODE" = "unrestricted" ]; then
  echo "Mode: unrestricted - all outbound traffic allowed"
  exit 0
fi

# Check if we have iptables capability
if ! command -v iptables >/dev/null 2>&1; then
  echo "Warning: iptables not available, cannot enforce network policy"
  exit 0
fi

# Test if we can actually use iptables (need NET_ADMIN capability)
if ! iptables -L OUTPUT >/dev/null 2>&1; then
  echo "Warning: Cannot access iptables (missing NET_ADMIN capability?)"
  echo "Network policy will NOT be enforced"
  exit 0
fi

echo "Configuring iptables rules..."

# Flush existing rules in OUTPUT chain (be careful not to lock ourselves out)
iptables -F OUTPUT 2>/dev/null || true

# Allow loopback
iptables -A OUTPUT -o lo -j ACCEPT

# Allow established connections
iptables -A OUTPUT -m state --state ESTABLISHED,RELATED -j ACCEPT

# Allow DNS (needed for resolution)
iptables -A OUTPUT -p udp --dport 53 -j ACCEPT
iptables -A OUTPUT -p tcp --dport 53 -j ACCEPT

echo "Adding essential domains..."
for domain in $ESSENTIAL_DOMAINS; do
  domain=$(echo "$domain" | xargs)  # trim whitespace
  if [ -n "$domain" ]; then
    allow_domain "$domain"
  fi
done

echo "Adding LLM provider domains..."
llm_domains=$(get_llm_domains)
for domain in $llm_domains; do
  domain=$(echo "$domain" | xargs)
  if [ -n "$domain" ]; then
    allow_domain "$domain"
  fi
done

# Add user allowlist domains if in allowlist mode
if [ "$NETWORK_MODE" = "allowlist" ] && [ -n "$NETWORK_ALLOWLIST" ]; then
  echo "Adding user allowlist domains..."
  echo "$NETWORK_ALLOWLIST" | tr ',' '\n' | while read domain; do
    domain=$(echo "$domain" | xargs)
    if [ -n "$domain" ]; then
      allow_domain "$domain"
    fi
  done
fi

# Drop everything else
echo "Setting default policy to DROP..."
iptables -A OUTPUT -j DROP

echo "=== Network policy applied ==="
echo ""
echo "Allowed destinations:"
iptables -L OUTPUT -n | grep ACCEPT | grep -v "state\|lo" | awk '{print "  - " $5}' | sort -u

