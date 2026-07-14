"""
Single source of truth for both tools' versions.

The two processors ship independently, so each has its own version, mirroring
the two-section structure of CHANGELOG.md. Bump the relevant constant, add a
CHANGELOG entry, commit, then push a matching tag (``ocean-vX.Y.Z`` /
``air-vX.Y.Z``). The GitHub Actions release workflow parses this file and
refuses to build if the tag's number disagrees with the constant here.
"""

OCEAN_VERSION = "1.7.0"
AIR_VERSION = "1.3.0"
