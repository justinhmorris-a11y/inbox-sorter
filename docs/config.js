/* Inbox Sorter - deployment settings. The client ID is not a secret. */
window.SORTER_CONFIG = {
  clientId: 'REPLACE-WITH-APPLICATION-CLIENT-ID',
  tenantId: 'REPLACE-WITH-DIRECTORY-TENANT-ID',
  // Rules to start with the very first time the add-in runs (after that the mailbox copy is used).
  // Codes: I keep in inbox, N Newsletters, R Receipts, T Notifications, D Delete, J Junk, F:<folder name>
  seedRules: {
    senders: {},
    domains: { 'golfbreaks.com': 'D', 'epsa.com': 'J' }
  }
};
