using ASCOM.Alpaca.Discovery;

namespace ConformU.Tests;

public class ExtensionMethodsTests
{
    [Fact]
    public void ToJsonNameCaseSensitivity_WhenTrue_ReturnsCorrectCasingOnly()
    {
        JsonNameCaseSensitivity result = true.ToJsonNameCaseSensitivity();

        Assert.Equal(JsonNameCaseSensitivity.CorrectCasingOnly, result);
    }

    [Fact]
    public void ToJsonNameCaseSensitivity_WhenFalse_ReturnsAnyCasing()
    {
        JsonNameCaseSensitivity result = false.ToJsonNameCaseSensitivity();

        Assert.Equal(JsonNameCaseSensitivity.AnyCasing, result);
    }
}
