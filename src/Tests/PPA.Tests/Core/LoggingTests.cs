using System;
using FluentAssertions;
using PPA.Logging;
using Xunit;

namespace PPA.Tests.Core
{
    /// <summary>
    /// 日志系统测试
    /// </summary>
    public class LoggingTests
    {
        [Fact]
        public void NullLogger_Should_Not_Throw()
        {
            // Arrange
            var logger = NullLogger.Instance;

            // Act & Assert - 所有方法都不应抛出异常
            var action = () =>
            {
                logger.LogInformation("Test message");
                logger.LogWarning("Test warning");
                logger.LogDebug("Test debug");
                logger.LogError("Test error");
                logger.LogError("Test error with exception", new Exception("Test"));
            };

            action.Should().NotThrow();
        }

        [Fact]
        public void NullLogger_Instance_Should_Be_Singleton()
        {
            // Arrange & Act
            var instance1 = NullLogger.Instance;
            var instance2 = NullLogger.Instance;

            // Assert
            instance1.Should().BeSameAs(instance2);
        }

        [Fact]
        public void ConsoleLogger_Should_Not_Throw()
        {
            // Arrange
            var logger = new ConsoleLogger();

            // Act & Assert
            var action = () =>
            {
                logger.LogInformation("Test message");
                logger.LogWarning("Test warning");
                logger.LogDebug("Test debug");
                logger.LogError("Test error");
            };

            action.Should().NotThrow();
        }

        [Fact]
        public void ConsoleLogger_Should_Respect_MinLevel()
        {
            // Arrange
            var logger = new ConsoleLogger(LogLevel.Warning);

            // Act & Assert - 不应抛出异常
            var action = () =>
            {
                logger.LogInformation("Should be ignored");
                logger.LogDebug("Should be ignored");
                logger.LogWarning("Should be logged");
                logger.LogError("Should be logged");
            };

            action.Should().NotThrow();
        }
    }
}
