// OneDrive sync settings, from the organisation's Entra ID app registration.
//
// Leave clientId empty and sync stays off: the app saves in this browser only,
// exactly as before. To turn it on, the admin registers the app (platform
// "Single-page application", redirect URI <where the app is hosted>/auth-redirect.html,
// delegated Microsoft Graph permission Files.ReadWrite.AppFolder) and puts the
// two IDs below. Neither is a secret.
//
// OneDrive names the folder after the registration's display name:
// OneDrive › Apps › <display name>.

export const SYNC_CONFIG = {
  clientId: "",          // Application (client) ID
  tenantId: "",          // Directory (tenant) ID
  scopes: ["Files.ReadWrite.AppFolder"],
};
