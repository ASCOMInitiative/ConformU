using ASCOM.Common;
using ASCOM.Common.DeviceInterfaces;
using System.Reflection;

namespace ConformU.Tests;

public class DeviceCapabilitySupportTests
{
    private static readonly Assembly ProductionAssembly = typeof(ConformConfiguration).Assembly;
    private static readonly Type CapabilitySupportType = ProductionAssembly.GetType(
        "ConformU.DeviceCapabilitySupport",
        throwOnError: true,
        ignoreCase: false)!;

    [Theory]
    [InlineData(0, false)]
    [InlineData(1, true)]
    [InlineData(2, false)]
    public void AlignerCapabilityMethods_ReturnExpectedV1State(int interfaceVersion, bool expected)
    {
        bool hasConnectAndDeviceState = InvokeBoolean(
            "HasConnectAndDeviceState",
            DeviceTypes.Aligner,
            interfaceVersion);
        bool isPlatform7OrLater = InvokeBoolean(
            "IsPlatform7OrLater",
            DeviceTypes.Aligner,
            interfaceVersion);

        Assert.Equal(expected, hasConnectAndDeviceState);
        Assert.Equal(expected, isPlatform7OrLater);
    }

    [Fact]
    public void CameraCapabilities_MatchLibraryImplementation()
    {
        const int interfaceVersion = 4;

        bool hasConnectAndDeviceState = InvokeBoolean(
            "HasConnectAndDeviceState",
            DeviceTypes.Camera,
            interfaceVersion);
        bool isPlatform7OrLater = InvokeBoolean(
            "IsPlatform7OrLater",
            DeviceTypes.Camera,
            interfaceVersion);

        Assert.Equal(
            DeviceCapabilities.HasConnectAndDeviceState(DeviceTypes.Camera, interfaceVersion),
            hasConnectAndDeviceState);
        Assert.Equal(
            DeviceCapabilities.IsPlatform7OrLater(DeviceTypes.Camera, interfaceVersion),
            isPlatform7OrLater);
    }

    [Fact]
    public void TryGetLatestPlatform6Interface_AlignerReturnsFalseWithZero()
    {
        MethodInfo method = CapabilitySupportType.GetMethod(
            "TryGetLatestPlatform6Interface",
            BindingFlags.NonPublic | BindingFlags.Static)!;
        object[] arguments = [DeviceTypes.Aligner, (short)0];

        bool found = (bool)method.Invoke(null, arguments)!;

        Assert.False(found);
        Assert.Equal((short)0, (short)arguments[1]);
    }

    [Fact]
    public void TryGetLatestPlatform6Interface_KnownDeviceMatchesLibrary()
    {
        DeviceTypes deviceType = DeviceTypes.Camera;
        Assert.True(DeviceCapabilities.LatestPlatform6Interface.TryGetValue(deviceType, out short libraryValue));

        MethodInfo method = CapabilitySupportType.GetMethod(
            "TryGetLatestPlatform6Interface",
            BindingFlags.NonPublic | BindingFlags.Static)!;
        object[] arguments = [deviceType, (short)0];

        bool found = (bool)method.Invoke(null, arguments)!;

        Assert.True(found);
        Assert.Equal(libraryValue, (short)arguments[1]);
    }

    private static bool InvokeBoolean(string methodName, DeviceTypes deviceType, int interfaceVersion)
    {
        MethodInfo method = CapabilitySupportType.GetMethod(
            methodName,
            BindingFlags.NonPublic | BindingFlags.Static)!;
        object[] arguments =
        [
            deviceType,
            interfaceVersion
        ];

        return (bool)method.Invoke(null, arguments)!;
    }
}
