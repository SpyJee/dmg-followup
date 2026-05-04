// Timer-triggered keep-warm for the dmg-followup-proxy Function App.
// Fires every 5 minutes. The act of executing IS the warm-up — by the
// time the worker would have scaled to zero on Consumption, this tick
// has already kept it alive, so the next user-facing HTTP request
// (followup-api / order-list / template-list / po-pdf, etc.) hits a
// warm worker instead of paying a 2-5s cold-start tax.
//
// Cost: ~17K executions/month at the Consumption rate of $0.20/million.
// First 1M executions/month are on the free tier, so this is effectively
// free.
module.exports = async function (context, myTimer) {
  context.log('keep-warm tick', new Date().toISOString());
};
