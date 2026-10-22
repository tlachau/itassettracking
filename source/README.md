# IT Asset Tracker - SPFx source

SharePoint Framework 1.23 web part. See the README one folder up for features and install steps.

    npm install
    npm run build      # -> sharepoint/solution/it-asset-tracker.sppkg

- `src/webparts/assetTracker/AssetTrackerWebPart.ts` - web part, property pane, mounts the app
- `src/webparts/assetTracker/app/app.ts` - the application (UI + SharePoint REST data layer)
- `src/webparts/assetTracker/app/appStyles.ts` - styles
