﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using Newtonsoft.Json;
using OfficeOpenXml;  // EPPlus Library

namespace AbbreviationWordAddin
{
    internal class AbbreviationManager
    {
        private static Dictionary<string, string> abbreviationDict = new Dictionary<string, string>();
        private static Dictionary<string, string> autoCorrectCache = new Dictionary<string, string>();

        private static bool isAutoCorrectCacheInitialized = false;
        private static string cacheFilePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "AbbreviationWordAddin",
            "abbreviations.json"
        );

        


        // Initialize AutoCorrect cache
        public static void InitializeAutoCorrectCache(Microsoft.Office.Interop.Word.AutoCorrect autoCorrect)
        {
            if (!isAutoCorrectCacheInitialized)
            {
                autoCorrectCache.Clear();
                for (int i = 1; i <= autoCorrect.Entries.Count; i++)
                {
                    string abbreviation = autoCorrect.Entries[i].Name;
                    string fullForm = autoCorrect.Entries[i].Value;
                    if (!string.IsNullOrEmpty(abbreviation) && !string.IsNullOrEmpty(fullForm))
                    {
                        autoCorrectCache[abbreviation] = fullForm;
                    }
                }

                //System.Windows.Forms.MessageBox.Show("Abbreviations loaded from AutoCorrect cache. Count: " + autoCorrectCache.Count.ToString());
                isAutoCorrectCacheInitialized = true;
            }
        }

        // Clear AutoCorrect cache
        public static void ClearAutoCorrectCache()
        {
            autoCorrectCache.Clear();
            isAutoCorrectCacheInitialized = false;
        }

        // Get replacement from cache
        public static string GetFromAutoCorrectCache(string text)
        {
            return autoCorrectCache.TryGetValue(text, out string replacement) ? replacement : null;
        }

        // Check if cache is initialized
        public static bool IsAutoCorrectCacheInitialized()
        {
            return isAutoCorrectCacheInitialized;
        }

        // Load abbreviations - first tries from cache, then from Excel if needed
        public static void LoadAbbreviations()
        {

            if (LoadFromCache())
            {
                return; // Successfully loaded from cache
            }

            // If cache doesn't exist or is invalid, load from Excel
            LoadFromExcel();
            SaveToCache(); // Save to cache for future use
        }

        public static void ClearCacheFile()
        {
            if (File.Exists(cacheFilePath))
            {
                File.Delete(cacheFilePath);
            }
        }

        // Load from local cache file
        private static bool LoadFromCache()
        {
            try
            {
                if (!File.Exists(cacheFilePath))
                    return false;

                string jsonContent = File.ReadAllText(cacheFilePath);
                var cache = JsonConvert.DeserializeObject<CacheData>(jsonContent);

                if (cache == null || cache.Version != Properties.Settings.Default.AbbreviationDataVersion || cache.Data == null)
                    return false;

                abbreviationDict = cache.Data;
                return true;
            }
            catch (Exception ex)
            {
                // If any error occurs during cache loading, we'll fall back to Excel
                System.Windows.Forms.MessageBox.Show(
                    "Failed to load abbreviations in autocorrect." + ex.Message,
                    "Please try restarting word.",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Warning
                   );
                
                Properties.Settings.Default.IsAutoCorrectLoaded = false;

                return false;
            }
        }

        // Save to local cache file
        private static void SaveToCache()
        {
            try
            {
                string directory = Path.GetDirectoryName(cacheFilePath);
                if (!Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                var cache = new CacheData
                {
                    Version = Properties.Settings.Default.AbbreviationDataVersion,
                    Data = abbreviationDict
                };

                string jsonContent = JsonConvert.SerializeObject(cache, Formatting.Indented);
                File.WriteAllText(cacheFilePath, jsonContent);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(
                    "Failed to save abbreviations cache: " + ex.Message,
                    "Cache Error",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Warning
                );
            }
        }

        class CacheData
        {
            public string Version { get; set; }
            public Dictionary<string, string> Data { get; set; }
        }

        // Load from embedded Excel file
        private static void LoadFromExcel()
        {
            try
            {
                AbbreviationManager.ClearAutoCorrectCache();

                ExcelPackage.License.SetNonCommercialPersonal("Abbreviations");
                var assembly = Assembly.GetExecutingAssembly();
                var resourceName = "AbbreviationWordAddin.Abbreviations.xlsx"; // Ensure the namespace matches your project

                using (Stream stream = assembly.GetManifestResourceStream(resourceName))

                {
                    //System.Windows.Forms.MessageBox.Show("Stream: " + stream);
                    if (stream == null)
                    {
                        throw new Exception("Excel file not found in embedded resources.");
                    }

                    using (var package = new ExcelPackage(stream))
                    {
                        //ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                        //System.Windows.Forms.MessageBox.Show("package: " + package);
                        ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // First sheet
                        int rowCount = worksheet.Dimension.Rows;

                        abbreviationDict.Clear(); // Clear existing entries
                        for (int row = 2; row <= rowCount; row++)  // Skip header row
                        {
                            string phrase = worksheet.Cells[row, 1].Text.Trim();
                            string abbreviation = worksheet.Cells[row, 2].Text.Trim();

                            if (!string.IsNullOrEmpty(phrase))
                            {
                                string lowerPhrase = phrase.ToLower(); // ✅ always lowercase!
                                if (!abbreviationDict.ContainsKey(lowerPhrase))
                                {
                                    abbreviationDict[lowerPhrase] = abbreviation;
                                }
                            }
                        }

                      

                       
                    }
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(
                    "Failed to load abbreviations from Excel: " + ex.Message,
                    "Abbreviation Load Error",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Warning
                );
            }
        }



        // Get abbreviation for a given phrase
        public static string GetAbbreviation(string phrase)
        {
            return abbreviationDict.ContainsKey(phrase) ? abbreviationDict[phrase] : phrase;
        }

        // Get all phrases for replacement
        public static List<string> GetAllPhrases()
        {
            return new List<string>(abbreviationDict.Keys)
                .OrderByDescending(p => p.Length)
                .ToList();
        }

        public static Dictionary<string, string> GetAllAbbreviations()
        {
            return abbreviationDict; 
        }
    }
}
