/* Inbox Sorter - deployment settings. The client ID is not a secret. */
window.SORTER_CONFIG = {
  clientId: '270f86ce-489d-48aa-a213-508ad54ee244',
  tenantId: 'de0c09d0-15b1-4606-a6ca-57408ee72392',
  // Rules to start with the very first time the add-in runs (after that the mailbox copy is used).
  // Codes: I keep in inbox, N Newsletters, R Receipts, T Notifications, D Delete, J Junk, F:<folder name>
  seedRules: {
    senders: {},
    domains: { 'golfbreaks.com': 'D', 'epsa.com': 'J' }
  }
};
