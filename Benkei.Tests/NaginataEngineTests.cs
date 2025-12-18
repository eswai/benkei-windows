using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Benkei;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Benkei.Tests
{
    [TestClass]
    public class NaginataEngineTests
    {
        [TestMethod]
        public void HandleKeyDownUp_Jkl_ProducesExpectedCharacters()
        {
            var engine = new NaginataEngine(LoadRulesFromYaml());

            var characters = new List<string>();

            characters.AddRange(ExtractOutputs(engine.HandleKeyDown(GetKeyCode("J"))));
            characters.AddRange(ExtractOutputs(engine.HandleKeyUp(GetKeyCode("J"))));

            characters.AddRange(ExtractOutputs(engine.HandleKeyDown(GetKeyCode("K"))));
            characters.AddRange(ExtractOutputs(engine.HandleKeyUp(GetKeyCode("K"))));

            characters.AddRange(ExtractOutputs(engine.HandleKeyDown(GetKeyCode("L"))));
            characters.AddRange(ExtractOutputs(engine.HandleKeyUp(GetKeyCode("L"))));
            
            Assert.AreEqual("AIU", string.Concat(characters));
        }

        private static IReadOnlyList<NaginataRule> LoadRulesFromYaml()
        {
            var configPath = Path.Combine(GetSolutionRoot(), "Naginata.yaml");
            Assert.IsTrue(File.Exists(configPath), $"Expected Naginata.yaml at '{configPath}'.");

            var loader = new NaginataConfigLoader();
            return loader.Load(configPath);
        }

        private static IEnumerable<string> ExtractOutputs(IEnumerable<NaginataAction> actions)
        {
            return actions?
                       .Where(action => action.Kind == NaginataActionType.Character || action.Kind == NaginataActionType.Tap)
                       .Select(action => action.Value)
                   ?? Enumerable.Empty<string>();
        }

        private static string GetSolutionRoot()
        {
            var baseDirectory = AppDomain.CurrentDomain.BaseDirectory;
            return Path.GetFullPath(Path.Combine(baseDirectory, "..", "..", "..", ".."));
        }

        private static int GetKeyCode(string keyName)
        {
            if (!VirtualKeyMapper.TryGetKeyCode(keyName, out var keyCode))
            {
                Assert.Fail($"Missing virtual key mapping for '{keyName}'.");
            }

            return keyCode;
        }
    }
}
