"use strict";

const { syncInvoicesFromSource } = require("./sync-invoices-from-source");

module.exports = {
  syncInvoicePreviewPayload: async (request = {}, context = {}) => {
    const baseRequest = request && typeof request === "object" ? request : {};
    const forcedRequest = Object.assign({}, baseRequest, {
      prefetchPreviewPdfs: true,
      searchOnly: false,
    });
    return syncInvoicesFromSource(forcedRequest, context);
  },
};

