/**
 * Provider-neutral billing contract.
 * Replace these methods with Stripe, Paddle, or another provider without
 * changing the HTTP routes that consume the interface.
 */
class BillingProvider {
  async createCheckout() {
    throw new Error("Billing provider is not configured");
  }

  async cancelSubscription() {
    throw new Error("Billing provider is not configured");
  }

  async getSubscription() {
    throw new Error("Billing provider is not configured");
  }
}

module.exports = { BillingProvider };
