const widgets = [
  {
    identifier: "invoice-access-manager",
    location: ["customer_portal_widget", "left_nav"],
    name: "Invoice Access Manager",
    description:
      "Import PDF invoices with role-based visibility for admins and PM-aligned collaborators.",
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
