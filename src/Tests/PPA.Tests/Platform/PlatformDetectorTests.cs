using FluentAssertions;
using PPA.Core.Abstraction;
using PPA.Universal.Platform;
using Xunit;

namespace PPA.Tests.Platform
{
    /// <summary>
    /// 平台检测器测试
    /// </summary>
    public class PlatformDetectorTests
    {
        [Fact]
        public void Detect_Should_Return_PlatformInfo()
        {
            // Act
            var info = PlatformDetector.Detect();

            // Assert
            info.Should().NotBeNull();
            info.ActivePlatform.Should().BeOneOf(
                PlatformType.Unknown,
                PlatformType.PowerPoint,
                PlatformType.WPS);
        }

        [Fact]
        public void Detect_Should_Cache_Result()
        {
            // Act
            var info1 = PlatformDetector.Detect();
            var info2 = PlatformDetector.Detect();

            // Assert
            info1.Should().BeSameAs(info2);
        }

        [Fact]
        public void Redetect_Should_Return_Fresh_Result()
        {
            // Arrange
            var info1 = PlatformDetector.Detect();

            // Act
            var info2 = PlatformDetector.Redetect();

            // Assert
            info2.Should().NotBeNull();
            // 新检测结果可能与缓存不同（不同对象实例）
        }

        [Fact]
        public void DetectFromApplication_Should_Return_Unknown_For_Null()
        {
            // Act
            var result = PlatformDetector.DetectFromApplication(null);

            // Assert
            result.Should().Be(PlatformType.Unknown);
        }

        [Fact]
        public void PlatformInfo_ToString_Should_Return_Description()
        {
            // Arrange
            var info = new PlatformInfo
            {
                PowerPointInstalled = true,
                PowerPointRunning = false,
                WPSInstalled = true,
                WPSRunning = false,
                ActivePlatform = PlatformType.PowerPoint
            };

            // Act
            var result = info.ToString();

            // Assert
            result.Should().Contain("PowerPoint");
            result.Should().Contain("WPS");
        }

        [Fact]
        public void PlatformInfo_HasAvailablePlatform_Should_Work()
        {
            // Arrange
            var infoNone = new PlatformInfo();
            var infoPPT = new PlatformInfo { PowerPointInstalled = true };
            var infoWPS = new PlatformInfo { WPSInstalled = true };

            // Assert
            infoNone.HasAvailablePlatform.Should().BeFalse();
            infoPPT.HasAvailablePlatform.Should().BeTrue();
            infoWPS.HasAvailablePlatform.Should().BeTrue();
        }
    }
}
