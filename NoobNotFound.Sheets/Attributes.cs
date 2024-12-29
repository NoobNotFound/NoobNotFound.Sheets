using System;

[AttributeUsage(AttributeTargets.Property)]
public class SheetColumnAttribute : Attribute
{
    public int Index { get; }

    public SheetColumnAttribute(int index)
    {
        Index = index;
    }
}

[AttributeUsage(AttributeTargets.Property)]
public class SheetIgnoreAttribute : Attribute
{
    public SheetIgnoreAttribute()
    {
    }
}
