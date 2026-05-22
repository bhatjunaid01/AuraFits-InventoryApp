# Aura Fits Tauri

Aura Fits migrated from Electron to Tauri + React + Vite with a local SQLite backend.

## Development

```powershell
npm.cmd install
npm.cmd run tauri:dev
```

Default admin login:

- Email: `admin@aura.fits`
- Password: `admin123`

## Production build

```powershell
npm.cmd run tauri:build
```

The SQLite database is created automatically in the app data directory as `aurafits.db`.
