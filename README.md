# SPFx Application Customizer – Conditional Banner

## Summary

An SharePoint Framework (SPFx) Application Customizer that renders a site-wide banner in the **Top** placeholder of modern SharePoint pages, and **conditionally hides** it when the user is viewing the `Employee` list.

The banner survives SharePoint's client-side (partial) navigation: it re-evaluates the current URL every time the user navigates or the placeholders change, so it appears and disappears correctly without a full page reload.

## Used SharePoint Framework Version

![version](https://img.shields.io/badge/version-1.23.0-green.svg)

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
- [Microsoft 365 tenant](https://docs.microsoft.com/sharepoint/dev/spfx/set-up-your-developer-tenant)

> Get your own free development tenant by subscribing to [Microsoft 365 developer program](http://aka.ms/o365devprogram)

## Prerequisites

- Node.js `>=22.14.0 <23.0.0`
- Heft (`npm install -g @rushstack/heft`)
- A SharePoint Online site with a list at `/Lists/Employee/` (or adjust the condition, see [Configuration](#configuration))

## Solution

| Solution                                | Component                                     |
| --------------------------------------- | --------------------------------------------- |
| SPFx-Application-Customizer-Conditional | `EmployeeRequestsBannerApplicationCustomizer` |

## Version history

| Version | Date             | Comments        |
| ------- | ---------------- | --------------- |
| 1.0     | October 30, 2026 | Initial release |

## Disclaimer

**THIS CODE IS PROVIDED _AS IS_ WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---

## Minimal Path to Awesome

1. Clone this repository and open the solution folder.
2. Install dependencies:

   ```bash
   npm install -g @rushstack/heft
   npm install
   ```

3. Update `pageUrl` in [config/serve.json](config/serve.json) to a page on your own tenant.
4. Start the local debug server:

   ```bash
   npm start        # heft start --clean
   ```

5. Open the page and accept the "Allow debug scripts" prompt.

### Build and deploy

```bash
npm run build    # heft test --clean --production && heft package-solution --production
```

Upload the generated `sharepoint/solution/sp-fx-application-customizer.sppkg` to your tenant or site collection App Catalog. Because `skipFeatureDeployment` is `true`, the extension can be deployed tenant-wide via the Tenant Wide Extensions list, or added per site.

## Features

- Adds a blue (`#0078d4`) banner to the **Top** placeholder on every modern page.
- Hides the banner on any URL containing `/lists/employee/` (case-insensitive).
- Re-renders on `navigatedEvent` and `placeholderProvider.changedEvent`, so it works with SharePoint's single-page navigation.
- Creates the banner only once per placeholder and recreates the placeholder if SharePoint disposes or detaches it.

This extension illustrates the following concepts:

- Using `PlaceholderName.Top` with `tryCreateContent`
- Handling client-side navigation in Application Customizers
- Conditional rendering based on the current page URL
- Passing configuration through `ClientSideComponentProperties`

## Configuration

### Properties

| Property        | Type     | Description                                                                          |
| --------------- | -------- | ------------------------------------------------------------------------------------ |
| `BannerMessage` | `string` | Text shown in the banner. Defaults to `Welcome to our SharePoint Portal!` when empty. |
| `testMessage`   | `string` | Scaffolded property; not used by the banner.                                         |

Set these in:

- [config/serve.json](config/serve.json) for local debugging
- [sharepoint/assets/elements.xml](sharepoint/assets/elements.xml) and [sharepoint/assets/ClientSideInstance.xml](sharepoint/assets/ClientSideInstance.xml) for deployment

For example:

```json
"properties": {
  "BannerMessage": "Submit your employee requests by Friday!"
}
```

> The configs currently only set `testMessage`, so the default banner text is shown until you add `BannerMessage`.

### Changing where the banner is hidden

Edit `_isEmployeeListPage()` in [EmployeeRequestsBannerApplicationCustomizer.ts](src/extensions/employeeRequestsBanner/EmployeeRequestsBannerApplicationCustomizer.ts) to match your list URL or add other conditions.

## References

- [Getting started with SharePoint Framework](https://docs.microsoft.com/sharepoint/dev/spfx/set-up-your-developer-tenant)
- [Build your first SharePoint Framework Extension](https://learn.microsoft.com/sharepoint/dev/spfx/extensions/get-started/build-a-hello-world-extension)
- [Use page placeholders from Application Customizer](https://learn.microsoft.com/sharepoint/dev/spfx/extensions/get-started/using-page-placeholder-with-extensions)
- [Tenant-wide deployment of extensions](https://learn.microsoft.com/sharepoint/dev/spfx/extensions/basics/tenant-wide-deployment-extensions)
- [Heft Documentation](https://heft.rushstack.io/)
