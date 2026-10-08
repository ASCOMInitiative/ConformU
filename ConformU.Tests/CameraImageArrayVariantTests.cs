using ASCOM;
using System.Reflection;

namespace ConformU.Tests;

public class CameraImageArrayVariantTests
{
    [Fact]
    public void AlpacaPropertyNotImplementedException_IsExpected()
    {
        Assert.True(IsExpectedAlpacaException(DeviceTechnology.Alpaca, new PropertyNotImplementedException()));
    }

    [Fact]
    public void AlpacaNotImplementedException_IsExpected()
    {
        Assert.True(IsExpectedAlpacaException(DeviceTechnology.Alpaca, new ASCOM.NotImplementedException()));
    }

    [Fact]
    public void UnrelatedExceptionAndComTechnology_AreNotExpected()
    {
        Assert.False(IsExpectedAlpacaException(DeviceTechnology.Alpaca, new System.InvalidOperationException()));
        Assert.False(IsExpectedAlpacaException(DeviceTechnology.COM, new PropertyNotImplementedException()));
    }

    [Fact]
    public void ReturnedVariantArray_MustHaveExpectedDimensionsAndObjectElements()
    {
        Assert.True(IsValidVariantArray(new object[4, 3], 4, 3));
        Assert.True(IsValidVariantArray(new object[4, 3, 2], 4, 3));
        Assert.False(IsValidVariantArray(new object[3, 4], 4, 3));
        Assert.False(IsValidVariantArray(new object[4, 3, 2], 3, 4));
        Assert.False(IsValidVariantArray(new int[4, 3], 4, 3));
        Assert.False(IsValidVariantArray(new object[4], 4, 3));
    }

    private static bool IsExpectedAlpacaException(DeviceTechnology technology, Exception exception) =>
        (bool)GetRule("IsExpectedAlpacaImageArrayVariantException").Invoke(null, [technology, exception])!;

    private static bool IsValidVariantArray(Array array, int x, int y) =>
        (bool)GetRule("IsValidImageArrayVariant").Invoke(null, [array, x, y])!;

    private static MethodInfo GetRule(string name) =>
        typeof(Program).Assembly.GetType("ConformU.CameraTester")!
            .GetMethod(name, BindingFlags.NonPublic | BindingFlags.Static)!;
}
