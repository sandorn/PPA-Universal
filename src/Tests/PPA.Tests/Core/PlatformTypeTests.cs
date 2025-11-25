using FluentAssertions;
using PPA.Core.Abstraction;
using Xunit;

namespace PPA.Tests.Core
{
    /// <summary>
    /// PlatformType 枚举测试
    /// </summary>
    public class PlatformTypeTests
    {
        [Fact]
        public void PlatformType_Should_Have_Expected_Values()
        {
            // Assert
            PlatformType.Unknown.Should().Be(PlatformType.Unknown);
            PlatformType.PowerPoint.Should().Be(PlatformType.PowerPoint);
            PlatformType.WPS.Should().Be(PlatformType.WPS);
        }

        [Fact]
        public void PlatformType_Unknown_Should_Be_Default()
        {
            // Arrange
            PlatformType defaultValue = default;

            // Assert
            defaultValue.Should().Be(PlatformType.Unknown);
        }
    }
}
