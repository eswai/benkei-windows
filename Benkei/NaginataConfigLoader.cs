using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using YamlDotNet.Serialization;
using YamlDotNet.Serialization.NamingConventions;

namespace Benkei
{
    internal sealed class NaginataConfigLoader
    {
        private readonly IDeserializer _deserializer = new DeserializerBuilder()
            .WithNamingConvention(UnderscoredNamingConvention.Instance)
            .IgnoreUnmatchedProperties()
            .Build();

        public IReadOnlyList<NaginataRule> Load(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
            {
                throw new ArgumentException("Config path is required.", nameof(filePath));
            }

            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException($"Naginata config not found at '{filePath}'.", filePath);
            }

            Console.WriteLine($"[Benkei] YAML読み込み開始: {filePath}");
            List<NaginataYamlCommand> commands;
            using (var reader = File.OpenText(filePath))
            {
                commands = _deserializer.Deserialize<List<NaginataYamlCommand>>(reader) ?? new List<NaginataYamlCommand>();
            }
            Console.WriteLine($"[Benkei] YAMLコマンド解析完了: {commands.Count}件");

            var rules = new List<NaginataRule>(commands.Count);
            foreach (var command in commands)
            {
                var mae = ToKeySet(command.Mae);
                var douji = ToKeySet(command.Douji);
                var actions = ToActions(command.Type);
                if (actions.Count == 0)
                {
                    continue;
                }

                rules.Add(new NaginataRule(mae, douji, actions));
            }

            Console.WriteLine($"[Benkei] ルール変換完了: {rules.Count}件");
            return rules;
        }

        private static HashSet<int> ToKeySet(IEnumerable<string> entries)
        {
            var set = new HashSet<int>();
            if (entries == null)
            {
                return set;
            }

            foreach (var entry in entries)
            {
                if (!VirtualKeyMapper.TryGetKeyCode(entry, out var keyCode))
                {
                    throw new InvalidDataException($"Unknown key '{entry}' in Naginata.yaml. Add the mapping in VirtualKeyMapper.");
                }

                set.Add(keyCode);
            }

            return set;
        }

        private static List<NaginataAction> ToActions(IEnumerable<Dictionary<string, string>> typeEntries)
        {
            var actions = new List<NaginataAction>();
            if (typeEntries == null)
            {
                return actions;
            }

            foreach (var entry in typeEntries)
            {
                if (entry == null || entry.Count == 0)
                {
                    continue;
                }

                foreach (var kvp in entry)
                {
                    var actionKey = kvp.Key.Trim().ToLowerInvariant();
                    NaginataActionType kind;
                    switch (actionKey)
                    {
                        case "tap":
                            kind = NaginataActionType.Tap;
                            break;
                        case "press":
                            kind = NaginataActionType.Press;
                            break;
                        case "release":
                            kind = NaginataActionType.Release;
                            break;
                        case "character":
                            kind = NaginataActionType.Character;
                            break;
                        case "repeat":
                            kind = NaginataActionType.Repeat;
                            break;
                        case "reset":
                            kind = NaginataActionType.Reset;
                            break;
                        default:
                            throw new InvalidDataException($"Unknown action '{kvp.Key}' in Naginata.yaml.");
                    }

                    actions.Add(new NaginataAction(kind, kvp.Value));
                }
            }

            return actions;
        }

        private sealed class NaginataYamlCommand
        {
            public List<string> Mae { get; set; } = new List<string>();

            public List<string> Douji { get; set; } = new List<string>();

            public List<Dictionary<string, string>> Type { get; set; } = new List<Dictionary<string, string>>();
        }
    }
}
