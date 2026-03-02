const widgets = [
  {
    identifier: "customer-home-bulletin-board",
    location: ["customer_portal_widget", "left_nav"],
    name: "Customer Bulletin Board",
    description:
      "Sticky-note bulletin board for customer updates and collaboration.",
    icon: "dist/icon.svg",
    entrypoint: {
      html: "dist/index.html",
    },
  },
];

module.exports = {
  version: "1.0.0",
  widgets,
  serverActions: [],
  eventHandlers: {},
  scheduledActions: [],
};
