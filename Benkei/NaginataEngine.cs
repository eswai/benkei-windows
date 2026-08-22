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
        public bool IsConversionKey(int keyCode) => VirtualKeyMapper.IsConversionKey(keyCode);

        public void Reset()
        {
            _pressedKeys.Clear();
            _pendingInput.Clear();
        }

        public List<NaginataAction> HandleKeyDown(int keyCode)
        {
            Logger.Log($"[Engine] KeyDown: {keyCode}, PressedKeys: [{string.Join(", ", _pressedKeys)}]");
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
                    Logger.Log($"[Engine] Candidate: [{string.Join(", ", candidate)}], Count: {candidateCount}");
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
            Logger.Log($"[Engine] PendingInput: {_pendingInput.Count}グループ");

            if (_pendingInput.Count > 1)
            {
                Logger.Log($"[Engine] Dequeue (複数グループ)");
                return DequeuePending();
            }

            if (_pendingInput.Count == 1 && NumberOfCandidates(_pendingInput[0]) == 1)
            {
                Logger.Log($"[Engine] Dequeue (確定)");
                return DequeuePending();
            }

            return new List<NaginataAction>();
        }

        public List<NaginataAction> HandleKeyUp(int keyCode)
        {
            Logger.Log($"[Engine] KeyUp: {keyCode}, PressedKeys: [{string.Join(", ", _pressedKeys)}]");
            _pressedKeys.Remove(keyCode);
            var result = new List<NaginataAction>();

            if (_pressedKeys.Count == 0)
            {
                Logger.Log($"[Engine] 全キー解放 - Pending処理: {_pendingInput.Count}グループ");
                while (_pendingInput.Count > 0)
                {
                    result.AddRange(NgType(_pendingInput[0]));
                    _pendingInput.RemoveAt(0);
                }
            }
            else
            {
                if (_pendingInput.Count > 0 && NumberOfCandidates(_pendingInput[0]) == 1)
                {
                    Logger.Log($"[Engine] 部分解放 - 確定処理");
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
                Logger.Log($"[Engine] NgType: 空入力");
                return new List<NaginataAction>();
            }

            Logger.Log($"[Engine] NgType: [{string.Join(", ", keys)}]");

            if (keys.Count == 1 && keys[0] == VirtualKeyMapper.Return)
            {
                Logger.Log($"[Engine] NgType: Returnキー検出");
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
                    Logger.Log($"[Engine] NgType: ルールマッチ - {actions.Count}個のアクション");
                    return actions;
                }
            }

            if (keys.Count > 1)
            {
                Logger.Log($"[Engine] NgType: 分割処理開始");
                var lead = keys.Take(keys.Count - 1).ToList();
                var lastKey = keys[keys.Count - 1];
                var tail = new List<int> { lastKey };
                var resolved = NgType(lead);
                resolved.AddRange(NgType(tail));
                return resolved;
            }

            Logger.Log($"[Engine] NgType: マッチなし");
            return new List<NaginataAction>();
        }

        private int NumberOfMatches(IList<int> keys)
        {
            return CountEntries(keys, false);
        }

        private int NumberOfCandidates(IList<int> keys)
        {
            return CountEntries(keys, true);
        }

        private int CountEntries(IList<int> keys, bool candidatesOnly)
        {
            if (keys == null || keys.Count == 0)
            {
                return 0;
            }

            var normalizedKeys = NormalizeKeys(keys);
            var key0 = new HashSet<int> { normalizedKeys[0] };
            var key1 = normalizedKeys.Count > 1
                ? new HashSet<int> { normalizedKeys[1] }
                : new HashSet<int>();
            var key2 = normalizedKeys.Count > 2
                ? new HashSet<int> { normalizedKeys[2] }
                : new HashSet<int>();
            var key01 = new HashSet<int>(key0);
            key01.UnionWith(key1);
            var key12 = new HashSet<int>(key1);
            key12.UnionWith(key2);
            var key012 = new HashSet<int>(key01);
            key012.UnionWith(key2);
            var count = 0;

            foreach (var rule in _rules)
            {
                bool matches;
                switch (normalizedKeys.Count)
                {
                    case 1:
                        matches = candidatesOnly
                            ? rule.Mae.IsSupersetOf(key0) || (rule.Mae.Count == 0 && rule.Douji.IsSupersetOf(key0))
                            : rule.Mae.SetEquals(key0) || (rule.Mae.Count == 0 && rule.Douji.SetEquals(key0));
                        break;
                    case 2:
                        matches = candidatesOnly
                            ? rule.Mae.SetEquals(key01)
                                || (rule.Mae.SetEquals(key0) && rule.Douji.IsSupersetOf(key1))
                                || (rule.Mae.Count == 0 && rule.Douji.IsSupersetOf(key01))
                            : rule.Mae.SetEquals(key01)
                                || (rule.Mae.SetEquals(key0) && rule.Douji.SetEquals(key1))
                                || (rule.Mae.Count == 0 && rule.Douji.SetEquals(key01));
                        break;
                    default:
                        matches = candidatesOnly
                            ? (rule.Mae.SetEquals(key01) && rule.Douji.IsSupersetOf(key2))
                                || (rule.Mae.SetEquals(key0) && rule.Douji.IsSupersetOf(key12))
                                || (rule.Mae.Count == 0 && rule.Douji.IsSupersetOf(key012))
                            : (rule.Mae.SetEquals(key01) && rule.Douji.SetEquals(key2))
                                || (rule.Mae.SetEquals(key0) && rule.Douji.SetEquals(key12))
                                || (rule.Mae.Count == 0 && rule.Douji.SetEquals(key012));
                        break;
                }

                if (!matches)
                {
                    continue;
                }

                if (candidatesOnly && normalizedKeys.Count == 2 && RuleHasMoreKeysThan(rule, key01))
                {
                    count = Math.Max(count, 2);
                }
                else
                {
                    count++;
                }

                if (count > 1)
                {
                    return 2;
                }
            }

            return count;
        }

        private static bool RuleHasMoreKeysThan(NaginataRule rule, HashSet<int> inputKeys)
        {
            var ruleKeys = new HashSet<int>(rule.Mae);
            ruleKeys.UnionWith(rule.Douji);
            return ruleKeys.Count > inputKeys.Count;
        }

        private static IList<int> NormalizeKeys(IList<int> keys)
        {
            if (keys == null)
            {
                return new List<int>();
            }

            return keys.Select(k => k == VirtualKeyMapper.Return ? VirtualKeyMapper.Space : k).ToList();
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
