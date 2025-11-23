using System.Collections.Generic;
using System.Linq;

namespace Benkei
{
    internal enum NaginataActionType
    {
        Tap,
        Press,
        Release,
        Character,
        Repeat,
        Reset
    }

    internal sealed class NaginataAction
    {
        public NaginataAction(NaginataActionType kind, string value)
        {
            Kind = kind;
            Value = value ?? string.Empty;
        }

        public NaginataActionType Kind { get; }

        public string Value { get; }
    }

    internal sealed class NaginataRule
    {
        public NaginataRule(HashSet<int> mae, HashSet<int> douji, IReadOnlyList<NaginataAction> actions)
        {
            Mae = mae;
            Douji = douji;
            Actions = actions;
        }

        public HashSet<int> Mae { get; }

        public HashSet<int> Douji { get; }

        public IReadOnlyList<NaginataAction> Actions { get; }

        public List<NaginataAction> CloneActions() => Actions.ToList();
    }
}
