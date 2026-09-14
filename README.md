# Loop Access Requests — Viva Connections Adaptive Card Extension

A SharePoint Framework (SPFx) **Adaptive Card Extension (ACE)** for the Viva Connections dashboard that gives approvers an at-a-glance view of pending **Microsoft Loop workspace access requests**, without leaving the dashboard.

- **Card view** shows a live count of pending requests.
- **Quick view** expands into an Adaptive Card listing each request with the requester's name and expiry date, sorted soonest-first.
- Data is read from a SharePoint list through the SPFx `SPHttpClient`, so no extra API permissions or backend are required.

![SPFx 1.23.2](https://img.shields.io/badge/SPFx-1.23.2-green.svg)
![Node 22](https://img.shields.io/badge/Node-22.x-blue.svg)
![TypeScript 5.8](https://img.shields.io/badge/TypeScript-5.8-blue.svg)
![Adaptive Cards 1.5](https://img.shields.io/badge/Adaptive%20Cards-1.5-purple.svg)

## Why

Loop workspace access requests typically land in email and get lost. Surfacing them on the Viva Connections dashboard, the home page employees already open on desktop and in Teams mobile, keeps the approval queue visible and shortens turnaround time.

## How it works

```
Viva Connections dashboard
        │
        ▼
LoopAccessRequestsAdaptiveCardExtension   (onInit → fetch, non-blocking)
        │  SPHttpClient GET
        │  /_api/web/lists/getbytitle('LoopAccessRequests')/items
        │    ?$select=Id,Title,ExpiresOn,Requester/Title
        │    &$expand=Requester
        │    &$filter=Status eq 'Pending'
        │    &$orderby=ExpiresOn asc
        ▼
   state { loading, requests[] }
        │
   ┌────┴─────────┐
   ▼              ▼
CardView      QuickView
"N pending    Adaptive Card template bound
 requests"    with ${requests} data + formatDateTime()
```

Key design points:

- **Non-blocking load.** `onInit` registers the views and kicks off the fetch without awaiting it, so the card renders immediately with a `Loading…` state and re-renders once data arrives via `setState`.
- **Declarative quick view.** The list UI is a pure Adaptive Card JSON template using `$data` binding and the `formatDateTime` expression, so the layout can be changed without touching TypeScript.
- **Lazy property pane.** The property pane bundle is loaded on demand with a dynamic `import()` and a named webpack chunk, keeping the runtime bundle small.
- **Typed state and props.** `ILoopRequest`, state, and props are all strongly typed and shared between the extension, card view, and quick view.

## Project structure

```
src/adaptiveCardExtensions/loopAccessRequests/
├── LoopAccessRequestsAdaptiveCardExtension.ts   # entry point, data fetch, view registration
├── LoopAccessRequestsAdaptiveCardExtension.manifest.json
├── LoopAccessRequestsPropertyPane.ts            # lazily loaded property pane (card title)
├── cardView/CardView.ts                         # compact dashboard card
├── quickView/QuickView.ts                       # binds state to the Adaptive Card template
├── quickView/template/QuickViewTemplate.json    # Adaptive Card 1.5 template
└── loc/                                         # localized strings
```

## Prerequisites

- Node.js 22.14 – 22.x (see `.nvmrc`)
- A Microsoft 365 tenant with SharePoint Online and Viva Connections
- A SharePoint list named **`LoopAccessRequests`** on the site where the dashboard lives, with these columns:

| Column      | Type                 | Notes                                  |
| ----------- | -------------------- | -------------------------------------- |
| `Title`     | Single line of text  | Loop workspace or page name            |
| `Requester` | Person or Group      | Who asked for access                   |
| `ExpiresOn` | Date and Time        | When the request lapses                |
| `Status`    | Choice               | Card shows only items set to `Pending` |

## Getting started

```bash
npm install
npm start          # heft start: builds and serves to the hosted workbench
```

Then open `https://<your-tenant>.sharepoint.com/_layouts/workbench.aspx`, add the **LoopAccessRequests** card, and it will read from the list on that site.

## Build and deploy

```bash
npm run build      # heft test --clean --production && heft package-solution --production
```

This produces `sharepoint/solution/loop-access-ace.sppkg`. Upload it to the tenant App Catalog, approve it, then add the card to a Viva Connections dashboard. The solution uses `skipFeatureDeployment: true`, so it is available tenant-wide without per-site installation.

## Tech stack

- [SharePoint Framework 1.23.2](https://aka.ms/spfx) with the Heft build rig (no Gulp)
- TypeScript 5.8, ESLint 9
- `@microsoft/sp-adaptive-card-extension-base` (`BaseComponentsCardView`, `BaseAdaptiveCardQuickView`)
- `@microsoft/sp-http` `SPHttpClient` for SharePoint REST
- Adaptive Cards 1.5 templating

## Possible next steps

- Approve or decline directly from the quick view using `Action.Execute` and a list item update.
- Make the list name and site URL configurable through the property pane.
- Add a "no pending requests" empty state and error state to the card view.
- Cache the last response so the card renders instantly on repeat visits.

## License

MIT
