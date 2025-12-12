using System;
using System.Collections.Generic;
using System.IO;
using YamlDotNet.Serialization;
using YamlDotNet.Serialization.NamingConventions;

namespace Benkei
{
    internal sealed class AlphabetConfigLoader
    {
        private readonly IDeserializer _deserializer = new DeserializerBuilder()
            .WithNamingConvention(UnderscoredNamingConvention.Instance)
            .IgnoreUnmatchedProperties()
            .Build();

        public AlphabetConfig Load(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
            {
                throw new ArgumentException("Config path is required.", nameof(filePath));
            }

            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException($"alphabet.yaml not found at '{filePath}'.", filePath);
            }

            Console.WriteLine($"[Benkei] alphabet.yaml 読み込み開始: {filePath}");

            AlphabetYaml yaml;
            using (var reader = File.OpenText(filePath))
            {
                yaml = _deserializer.Deserialize<AlphabetYaml>(reader) ?? new AlphabetYaml();
            }

            var kanaOnKeys = new HashSet<int>();
            if (yaml.KanaOn != null)
            {
                foreach (var entry in yaml.KanaOn)
                {
                    if (!VirtualKeyMapper.TryGetKeyCode(entry, out var keyCode))
                    {
                        throw new InvalidDataException($"Unknown key '{entry}' in alphabet.yaml kana_on.");
                    }

                    kanaOnKeys.Add(keyCode);
                }
            }

            var remap = new Dictionary<int, int>();
            if (yaml.Remap != null)
            {
                foreach (var kvp in yaml.Remap)
                {
                    if (!VirtualKeyMapper.TryGetKeyCode(kvp.Key, out var from))
                    {
                        throw new InvalidDataException($"Unknown key '{kvp.Key}' in alphabet.yaml remap.");
                    }

                    if (!VirtualKeyMapper.TryGetKeyCode(kvp.Value, out var to))
                    {
                        throw new InvalidDataException($"Unknown key '{kvp.Value}' in alphabet.yaml remap.");
                    }

                    remap[from] = to;
                }
            }

            Console.WriteLine($"[Benkei] alphabet.yaml remap数: {remap.Count}");
            return new AlphabetConfig(kanaOnKeys, remap);
        }

        private sealed class AlphabetYaml
        {
            public List<string> KanaOn { get; set; } = new List<string>();

            public Dictionary<string, string> Remap { get; set; } = new Dictionary<string, string>();
        }
    }

    internal sealed class AlphabetConfig
    {
        public AlphabetConfig(HashSet<int> kanaOn, Dictionary<int, int> remap)
        {
            KanaOn = kanaOn ?? new HashSet<int>();
            Remap = remap ?? new Dictionary<int, int>();
        }

        public IReadOnlyCollection<int> KanaOn { get; }

        public IReadOnlyDictionary<int, int> Remap { get; }

        public bool TryGetRemappedKey(int keyCode, out int mapped) => Remap.TryGetValue(keyCode, out mapped);
    }
}
