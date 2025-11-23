using System;
using System.Collections.Generic;
using System.Linq;

namespace Benkei
{
    internal sealed class NaginataEngine
    {
        private static readonly List<int[]> ShiftChains = new List<int[]>
        {
            Chain("D", "F"),
            Chain("C", "V"),
            Chain("J", "K"),
            Chain("M", "Comma"),
            Chain("Space"),
            Chain("Return"),
            Chain("F"),
            Chain("V"),
            Chain("J"),
            Chain("M")
        };

        private readonly List<NaginataRule> _rules;
        private readonly HashSet<int> _pressedKeys = new HashSet<int>();
        private readonly List<List<int>> _pendingInput = new List<List<int>>();

        public NaginataEngine(IReadOnlyList<NaginataRule> rules)
        {
            _rules = rules?.ToList() ?? throw new ArgumentNullException(nameof(rules));
            if (_rules.Count == 0)
            {
                throw new ArgumentException("At least one rule must be provided.", nameof(rules));
            }
        }

        public bool IsNaginataKey(int keyCode) => VirtualKeyMapper.IsNaginataKey(keyCode);

        public void Reset()
        {
            _pressedKeys.Clear();
            _pendingInput.Clear();
        }

        public List<NaginataAction> HandleKeyDown(int keyCode)
        {
            Console.WriteLine($"[Engine] KeyDown: {keyCode}, PressedKeys: [{string.Join(", ", _pressedKeys)}]");
            _pressedKeys.Add(keyCode);

            if (keyCode == VirtualKeyMapper.Space || keyCode == VirtualKeyMapper.Return)
            {
                _pendingInput.Add(new List<int> { keyCode });
            }
            else if (_pendingInput.Count > 0)
            {
                var last = _pendingInput[_pendingInput.Count - 1];
                if (last.Count > 0 && last[last.Count - 1] != keyCode)
                {
                    var candidate = last.Concat(new[] { keyCode }).ToList();
                    var candidateCount = NumberOfCandidates(candidate);
                    Console.WriteLine($"[Engine] Candidate: [{string.Join(", ", candidate)}], Count: {candidateCount}");
                    if (candidateCount > 0)
                    {
                        _pendingInput[_pendingInput.Count - 1] = candidate;
                    }
                    else
                    {
                        _pendingInput.Add(new List<int> { keyCode });
                    }
                }
                else
                {
                    _pendingInput.Add(new List<int> { keyCode });
                }
            }
            else
            {
                _pendingInput.Add(new List<int> { keyCode });
            }

            ApplyShiftChains(keyCode);
            Console.WriteLine($"[Engine] PendingInput: {_pendingInput.Count}グループ");

            if (_pendingInput.Count > 1)
            {
                Console.WriteLine($"[Engine] Dequeue (複数グループ)");
                return DequeuePending();
            }

            if (_pendingInput.Count == 1 && NumberOfCandidates(_pendingInput[0]) == 1)
            {
                Console.WriteLine($"[Engine] Dequeue (確定)");
                return DequeuePending();
            }

            return new List<NaginataAction>();
        }

        public List<NaginataAction> HandleKeyUp(int keyCode)
        {
            Console.WriteLine($"[Engine] KeyUp: {keyCode}, PressedKeys: [{string.Join(", ", _pressedKeys)}]");
            _pressedKeys.Remove(keyCode);
            var result = new List<NaginataAction>();

            if (_pressedKeys.Count == 0)
            {
                Console.WriteLine($"[Engine] 全キー解放 - Pending処理: {_pendingInput.Count}グループ");
                while (_pendingInput.Count > 0)
                {
                    result.AddRange(NgType(_pendingInput[0]));
                    _pendingInput.RemoveAt(0);
                }
            }
            else
            {
                _pendingInput.Add(new List<int>());
                if (_pendingInput.Count > 0 && NumberOfCandidates(_pendingInput[0]) == 1)
                {
                    Console.WriteLine($"[Engine] 部分解放 - 確定処理");
                    result.AddRange(NgType(_pendingInput[0]));
                    _pendingInput.RemoveAt(0);
                }
            }

            return result;
        }

        private List<NaginataAction> DequeuePending()
        {
            if (_pendingInput.Count == 0)
            {
                return new List<NaginataAction>();
            }

            var target = _pendingInput[0];
            _pendingInput.RemoveAt(0);
            return NgType(target);
        }

        private void ApplyShiftChains(int latestKey)
        {
            if (_pendingInput.Count == 0)
            {
                return;
            }

            var lastGroup = _pendingInput[_pendingInput.Count - 1];
            if (lastGroup.Count == 0)
            {
                return;
            }

            foreach (var chain in ShiftChains)
            {
                if (chain.Length == 0)
                {
                    continue;
                }

                if (chain.Contains(latestKey))
                {
                    continue;
                }

                if (chain.All(code => lastGroup.Contains(code)))
                {
                    continue;
                }

                if (!chain.All(_pressedKeys.Contains))
                {
                    continue;
                }

                var merged = MergeChain(chain, lastGroup);
                if (NumberOfMatches(merged) > 0)
                {
                    _pendingInput[_pendingInput.Count - 1] = merged;
                    break;
                }
            }
        }

        private List<NaginataAction> NgType(List<int> keys)
        {
            if (keys == null || keys.Count == 0)
            {
                Console.WriteLine($"[Engine] NgType: 空入力");
                return new List<NaginataAction>();
            }

            Console.WriteLine($"[Engine] NgType: [{string.Join(", ", keys)}]");

            if (keys.Count == 1 && keys[0] == VirtualKeyMapper.Return)
            {
                Console.WriteLine($"[Engine] NgType: Returnキー検出");
                return new List<NaginataAction> { new NaginataAction(NaginataActionType.Tap, "Return") };
            }

            var normalized = new HashSet<int>(keys.Select(k => k == VirtualKeyMapper.Return ? VirtualKeyMapper.Space : k));

            foreach (var rule in _rules)
            {
                var combined = new HashSet<int>(rule.Mae);
                combined.UnionWith(rule.Douji);
                if (combined.SetEquals(normalized))
                {
                    var actions = rule.CloneActions();
                    Console.WriteLine($"[Engine] NgType: ルールマッチ - {actions.Count}個のアクション");
                    return actions;
                }
            }

            if (keys.Count > 1)
            {
                Console.WriteLine($"[Engine] NgType: 分割処理開始");
                var lead = keys.Take(keys.Count - 1).ToList();
                var lastKey = keys[keys.Count - 1];
                var tail = new List<int> { lastKey };
                var resolved = NgType(lead);
                resolved.AddRange(NgType(tail));
                return resolved;
            }

            Console.WriteLine($"[Engine] NgType: マッチなし");
            return new List<NaginataAction>();
        }

        private int NumberOfMatches(IList<int> keys)
        {
            if (keys.Count == 0)
            {
                return 0;
            }

            var keySet = new HashSet<int>(keys);
            var matches = 0;

            foreach (var rule in _rules)
            {
                switch (keys.Count)
                {
                    case 1:
                        if (rule.Mae.SetEquals(keySet))
                        {
                            matches++;
                        }

                        if (rule.Mae.Count == 0 && rule.Douji.SetEquals(keySet))
                        {
                            matches++;
                        }

                        break;
                    case 2:
                        if (rule.Mae.SetEquals(keySet))
                        {
                            matches++;
                        }

                        if (rule.Mae.SetEquals(SliceSet(keys, 0, 1)) && rule.Douji.SetEquals(SliceSet(keys, 1)))
                        {
                            matches++;
                        }

                        if (rule.Mae.Count == 0 && rule.Douji.SetEquals(keySet))
                        {
                            matches++;
                        }

                        break;
                    default:
                        if (rule.Mae.SetEquals(SliceSet(keys, 0, 2)) && rule.Douji.SetEquals(SliceSet(keys, 2)))
                        {
                            matches++;
                        }

                        if (rule.Mae.SetEquals(SliceSet(keys, 0, 1)) && rule.Douji.SetEquals(SliceSet(keys, 1)))
                        {
                            matches++;
                        }

                        if (rule.Mae.Count == 0 && rule.Douji.SetEquals(keySet))
                        {
                            matches++;
                        }

                        break;
                }
            }

            return matches;
        }

        private int NumberOfCandidates(IList<int> keys)
        {
            if (keys.Count == 0)
            {
                return 0;
            }

            var keySet = new HashSet<int>(keys);
            var candidates = 0;

            foreach (var rule in _rules)
            {
                switch (keys.Count)
                {
                    case 1:
                        if (rule.Mae.IsSupersetOf(keySet))
                        {
                            candidates++;
                        }

                        if (rule.Mae.Count == 0 && rule.Douji.IsSupersetOf(keySet))
                        {
                            candidates++;
                        }

                        break;
                    case 2:
                        if (rule.Mae.SetEquals(keySet))
                        {
                            candidates++;
                        }

                        if (rule.Mae.SetEquals(SliceSet(keys, 0, 1)) && rule.Douji.IsSupersetOf(SliceSet(keys, 1)))
                        {
                            candidates++;
                        }

                        if (rule.Mae.Count == 0 && rule.Douji.IsSupersetOf(keySet))
                        {
                            candidates = rule.Douji.Count > 2 ? Math.Max(candidates, 2) : candidates + 1;
                        }

                        break;
                    default:
                        if (rule.Mae.SetEquals(SliceSet(keys, 0, 2)) && rule.Douji.IsSupersetOf(SliceSet(keys, 2)))
                        {
                            candidates++;
                        }

                        if (rule.Mae.SetEquals(SliceSet(keys, 0, 1)) && rule.Douji.IsSupersetOf(SliceSet(keys, 1)))
                        {
                            candidates++;
                        }

                        if (rule.Mae.Count == 0 && rule.Douji.IsSupersetOf(keySet))
                        {
                            candidates++;
                        }

                        break;
                }
            }

            return candidates;
        }

        private static HashSet<int> SliceSet(IList<int> source, int startIndex, int length)
        {
            var set = new HashSet<int>();
            if (source == null || startIndex >= source.Count || length <= 0)
            {
                return set;
            }

            var maxLength = Math.Min(length, source.Count - startIndex);
            for (var i = 0; i < maxLength; i++)
            {
                set.Add(source[startIndex + i]);
            }

            return set;
        }

        private static HashSet<int> SliceSet(IList<int> source, int startIndex)
        {
            return SliceSet(source, startIndex, source.Count - startIndex);
        }

        private static List<int> MergeChain(IEnumerable<int> chain, IEnumerable<int> existing)
        {
            var merged = new List<int>();
            foreach (var code in chain)
            {
                if (!merged.Contains(code))
                {
                    merged.Add(code);
                }
            }

            foreach (var code in existing)
            {
                if (!merged.Contains(code))
                {
                    merged.Add(code);
                }
            }

            return merged;
        }

        private static int[] Chain(params string[] keys)
        {
            return keys.Select(k =>
            {
                if (!VirtualKeyMapper.TryGetKeyCode(k, out var code))
                {
                    throw new InvalidOperationException($"Missing shift chain key mapping for '{k}'.");
                }

                return code;
            }).ToArray();
        }
    }
}
