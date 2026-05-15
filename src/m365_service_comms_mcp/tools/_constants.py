"""Constants shared across the v0.1 tools."""

from __future__ import annotations

# Per-tool maximum on the Graph ``$top`` parameter. We cap below Graph's own
# soft limit (1000 for messages, 500 for health overviews) so a single tool
# call cannot blow up the LLM's context window or the Graph rate budget.
MAX_TOP = 50
