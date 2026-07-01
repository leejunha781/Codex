using System;
using System.Collections.Generic;
using UnityEngine;

namespace SaveSystem
{
    /// <summary>
    /// Serializable container for all persistent game state.
    /// Add fields here as the game grows; only plain data types
    /// (primitives, strings, and other [Serializable] types) survive JSON serialization.
    /// </summary>
    [Serializable]
    public class GameData
    {
        [Tooltip("Schema version. Bump when the data layout changes so migrations can be applied on load.")]
        public int version = 1;

        public string playerName = "Player";
        public int level = 1;
        public int coins = 0;
        public float playtimeSeconds = 0f;

        public Vector3 lastPosition = Vector3.zero;

        public List<string> unlockedItems = new List<string>();

        // Unix timestamp (UTC seconds) of the last save. Set automatically by SaveManager.
        public long savedAtUnixSeconds = 0;
    }
}
