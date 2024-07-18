
namespace Config
{
    public struct ClassParam(string name, ClassType type)
    {
        public readonly string Name = name;
        public readonly ClassType Type = type;
    }

    public enum ClassType
    {
        List,
        Dict,
    }
}