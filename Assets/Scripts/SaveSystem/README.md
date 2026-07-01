# Save System (데이터 저장)

A lightweight, dependency-free save/load system for Unity that persists game
state to disk as JSON under `Application.persistentDataPath`.

## Files

- `GameData.cs` — the serializable data model. Add your persistent fields here.
- `SaveManager.cs` — static API for saving, loading, and deleting save files.
- `SaveManagerExample.cs` — a `MonoBehaviour` demonstrating usage.

## Usage

```csharp
using SaveSystem;

// Save
var data = new GameData { playerName = "Neo", level = 5, coins = 120 };
SaveManager.Save(data);

// Load (returns fresh GameData if no save exists)
GameData loaded = SaveManager.Load();

// Load only if a save exists
if (SaveManager.TryLoad(out GameData existing))
{
    Debug.Log($"Welcome back, {existing.playerName}");
}

// Delete
SaveManager.Delete();
```

## Design notes

- **Atomic writes:** saves are written to a `.tmp` file and then swapped into
  place, so an interrupted write cannot corrupt an existing save.
- **Safe loads:** malformed or missing files return a fresh `GameData` instead
  of throwing.
- **Versioning:** `GameData.version` is included so future schema changes can be
  migrated on load.
- **Multiple slots:** every method accepts an optional `fileName`, so you can
  keep separate save slots (e.g. `SaveManager.Save(data, "slot2.json")`).

## Where saves live

`SaveManager.GetSavePath()` returns the full path. Typical locations:

- Windows: `%userprofile%\AppData\LocalLow\<company>\<product>\savegame.json`
- macOS: `~/Library/Application Support/<company>/<product>/savegame.json`
- Linux: `~/.config/unity3d/<company>/<product>/savegame.json`
