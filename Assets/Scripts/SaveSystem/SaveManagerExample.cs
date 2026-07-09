using UnityEngine;

namespace SaveSystem
{
    /// <summary>
    /// Minimal example showing how to use <see cref="SaveManager"/>.
    /// Attach to a GameObject and use the context menu or hotkeys in Play mode.
    /// </summary>
    public class SaveManagerExample : MonoBehaviour
    {
        [SerializeField] private KeyCode saveKey = KeyCode.F5;
        [SerializeField] private KeyCode loadKey = KeyCode.F9;

        private GameData _data = new GameData();

        private void Awake()
        {
            _data = SaveManager.Load();
            Debug.Log($"[SaveManagerExample] Loaded data. Coins={_data.coins}, Level={_data.level}");
        }

        private void Update()
        {
            _data.playtimeSeconds += Time.deltaTime;

            if (Input.GetKeyDown(saveKey))
            {
                SaveGame();
            }
            else if (Input.GetKeyDown(loadKey))
            {
                _data = SaveManager.Load();
                Debug.Log($"[SaveManagerExample] Reloaded. Coins={_data.coins}, Level={_data.level}");
            }
        }

        [ContextMenu("Save Game")]
        private void SaveGame()
        {
            bool ok = SaveManager.Save(_data);
            Debug.Log(ok
                ? $"[SaveManagerExample] Saved to {SaveManager.GetSavePath()}"
                : "[SaveManagerExample] Save failed.");
        }

        [ContextMenu("Delete Save")]
        private void DeleteSave()
        {
            bool deleted = SaveManager.Delete();
            Debug.Log(deleted
                ? "[SaveManagerExample] Save deleted."
                : "[SaveManagerExample] No save file to delete.");
        }
    }
}
