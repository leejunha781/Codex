using System;
using System.IO;
using UnityEngine;

namespace SaveSystem
{
    /// <summary>
    /// Handles persisting <see cref="GameData"/> to disk as JSON under
    /// <see cref="Application.persistentDataPath"/>.
    ///
    /// Writes are performed atomically (write to a temp file, then replace)
    /// so a crash mid-write cannot corrupt an existing save.
    /// </summary>
    public static class SaveManager
    {
        private const string DefaultFileName = "savegame.json";

        public static string SaveDirectory => Application.persistentDataPath;

        public static string GetSavePath(string fileName = DefaultFileName)
        {
            return Path.Combine(SaveDirectory, fileName);
        }

        public static bool SaveExists(string fileName = DefaultFileName)
        {
            return File.Exists(GetSavePath(fileName));
        }

        /// <summary>
        /// Serializes <paramref name="data"/> to JSON and writes it to disk atomically.
        /// Returns true on success.
        /// </summary>
        public static bool Save(GameData data, string fileName = DefaultFileName, bool prettyPrint = true)
        {
            if (data == null)
            {
                Debug.LogError("[SaveManager] Cannot save null GameData.");
                return false;
            }

            data.savedAtUnixSeconds = DateTimeOffset.UtcNow.ToUnixTimeSeconds();

            string path = GetSavePath(fileName);
            string tempPath = path + ".tmp";

            try
            {
                string json = JsonUtility.ToJson(data, prettyPrint);

                Directory.CreateDirectory(SaveDirectory);
                File.WriteAllText(tempPath, json);

                // Atomic replace: File.Replace requires the destination to exist.
                if (File.Exists(path))
                {
                    File.Replace(tempPath, path, null);
                }
                else
                {
                    File.Move(tempPath, path);
                }

                return true;
            }
            catch (Exception e)
            {
                Debug.LogError($"[SaveManager] Failed to save to '{path}': {e}");
                CleanupTempFile(tempPath);
                return false;
            }
        }

        /// <summary>
        /// Loads and deserializes <see cref="GameData"/> from disk.
        /// Returns a fresh <see cref="GameData"/> instance if no save exists or loading fails.
        /// </summary>
        public static GameData Load(string fileName = DefaultFileName)
        {
            string path = GetSavePath(fileName);

            if (!File.Exists(path))
            {
                return new GameData();
            }

            try
            {
                string json = File.ReadAllText(path);
                GameData data = JsonUtility.FromJson<GameData>(json);

                if (data == null)
                {
                    Debug.LogWarning($"[SaveManager] Save at '{path}' deserialized to null. Returning fresh data.");
                    return new GameData();
                }

                return data;
            }
            catch (Exception e)
            {
                Debug.LogError($"[SaveManager] Failed to load from '{path}': {e}");
                return new GameData();
            }
        }

        /// <summary>
        /// Attempts to load existing save data. Returns false (and default data) when no valid save exists.
        /// </summary>
        public static bool TryLoad(out GameData data, string fileName = DefaultFileName)
        {
            if (!SaveExists(fileName))
            {
                data = new GameData();
                return false;
            }

            data = Load(fileName);
            return true;
        }

        /// <summary>
        /// Deletes the save file if it exists. Returns true if a file was deleted.
        /// </summary>
        public static bool Delete(string fileName = DefaultFileName)
        {
            string path = GetSavePath(fileName);
            try
            {
                if (File.Exists(path))
                {
                    File.Delete(path);
                    return true;
                }
            }
            catch (Exception e)
            {
                Debug.LogError($"[SaveManager] Failed to delete '{path}': {e}");
            }

            return false;
        }

        private static void CleanupTempFile(string tempPath)
        {
            try
            {
                if (File.Exists(tempPath))
                {
                    File.Delete(tempPath);
                }
            }
            catch (Exception e)
            {
                Debug.LogWarning($"[SaveManager] Failed to clean up temp file '{tempPath}': {e}");
            }
        }
    }
}
