# Full Change Log

- [v1.1.0]
  - Feat: Added support for hashtables
  - Feat: Added a 'ResetKey' switch to remove stale data from existing target keys prior to writing to them
  - Feat: Changed the 'EnsureUniqueness' switch to 'AllowOverwrite' and flipped the logic accordingly
  - Fix: Resolved a clobber condition for the $KeyName parameter when using -UseFirstPropertyAsKey
  - Fix: Implemented a hashset to add numeric counters for non-unique names in a manner which increments per-name
  - Fix: Implemented a less fragile approach for $DValue construction (which calculates leading zeros)
  - Fix: Added logic to ensure homogenious collections (no mixed PSObject and HashTables)
  - Docs: Updated documentation to reflect new switches, switch renames and new inputs/examples for hashtables
  - Test: Added pester tests for most permutations of input types, hives, cardinality, and switch combinations (Requires Pester 5.x+)

- [v1.0.1]
  - Fix: Renamed 'RootKeyName' to 'KeyName' within comment-based help section

- [v1.0.0]
  - Initial Release