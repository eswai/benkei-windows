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
            var keyEvents = new[]
            {
                (KeyCode: GetKeyCode("J"), IsKeyUp: false),
                (KeyCode: GetKeyCode("J"), IsKeyUp: true),
                (KeyCode: GetKeyCode("K"), IsKeyUp: false),
                (KeyCode: GetKeyCode("K"), IsKeyUp: true),
                (KeyCode: GetKeyCode("L"), IsKeyUp: false),
                (KeyCode: GetKeyCode("L"), IsKeyUp: true)
            };

            foreach (var (keyCode, isKeyUp) in keyEvents)
            {
                var actions = isKeyUp ? engine.HandleKeyUp(keyCode) : engine.HandleKeyDown(keyCode);
                characters.AddRange(ExtractOutputs(actions));
            }
            
            Assert.AreEqual("AIU", string.Concat(characters));
        }

        [TestMethod]
        public void ImeOffSingleLetterIsPassedThrough()
        {
            var engine = new NaginataEngine(LoadRulesFromYaml("alphabet.yaml"));
            var events = new[]
            {
                (KeyName: "Q", IsKeyUp: false),
                (KeyName: "Q", IsKeyUp: true)
            };

            Assert.AreEqual("Q", RunEvents(engine, events));
        }

        [TestMethod]
        public void V18ShiftedKanaAndThreeKeyShiftAreAvailable()
        {
            var engine = new NaginataEngine(LoadRulesFromYaml());
            var shiftedKana = new[]
            {
                (KeyName: "Space", IsKeyUp: false),
                (KeyName: "W", IsKeyUp: false),
                (KeyName: "W", IsKeyUp: true),
                (KeyName: "Space", IsKeyUp: true)
            };
            var shiftedEnter = new[]
            {
                (KeyName: "Space", IsKeyUp: false),
                (KeyName: "V", IsKeyUp: false),
                (KeyName: "M", IsKeyUp: false),
                (KeyName: "M", IsKeyUp: true),
                (KeyName: "V", IsKeyUp: true),
                (KeyName: "Space", IsKeyUp: true)
            };

            Assert.AreEqual("NE", RunEvents(engine, shiftedKana));
            Assert.AreEqual("Return", RunEvents(engine, shiftedEnter));
        }

        private static IReadOnlyList<NaginataRule> LoadRulesFromYaml(string fileName = "Naginata.yaml")
        {
            var configPath = Path.Combine(GetSolutionRoot(), fileName);
            Assert.IsTrue(File.Exists(configPath), $"Expected Naginata.yaml at '{configPath}'.");

            var loader = new NaginataConfigLoader();
            return loader.Load(configPath);
        }

        private static string RunEvents(NaginataEngine engine, IEnumerable<(string KeyName, bool IsKeyUp)> events)
        {
            var outputs = new List<string>();
            foreach (var keyEvent in events)
            {
                var keyCode = GetKeyCode(keyEvent.KeyName);
                var actions = keyEvent.IsKeyUp
                    ? engine.HandleKeyUp(keyCode)
                    : engine.HandleKeyDown(keyCode);
                outputs.AddRange(ExtractOutputs(actions));
            }

            return string.Concat(outputs);
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
