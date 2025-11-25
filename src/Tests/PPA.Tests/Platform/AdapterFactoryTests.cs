using System;
using FluentAssertions;
using Microsoft.Extensions.DependencyInjection;
using PPA.Core.Abstraction;
using PPA.Universal.Platform;
using Xunit;

namespace PPA.Tests.Platform
{
    /// <summary>
    /// 适配器工厂测试
    /// </summary>
    public class AdapterFactoryTests
    {
        [Fact]
        public void RegisterAdapter_PowerPoint_Should_Register_Services()
        {
            // Arrange
            var factory = new AdapterFactory();
            var services = new ServiceCollection();

            // Act
            factory.RegisterAdapter(services, PlatformType.PowerPoint);
            var provider = services.BuildServiceProvider();

            // Assert
            provider.GetService<IShapeOperations>().Should().NotBeNull();
            provider.GetService<ITableOperations>().Should().NotBeNull();
            provider.GetService<ISlideOperations>().Should().NotBeNull();
        }

        [Fact]
        public void RegisterAdapter_WPS_Should_Register_Services()
        {
            // Arrange
            var factory = new AdapterFactory();
            var services = new ServiceCollection();

            // Act
            factory.RegisterAdapter(services, PlatformType.WPS);
            var provider = services.BuildServiceProvider();

            // Assert
            provider.GetService<IShapeOperations>().Should().NotBeNull();
            provider.GetService<ITableOperations>().Should().NotBeNull();
            provider.GetService<ISlideOperations>().Should().NotBeNull();
        }

        [Fact]
        public void RegisterAdapter_Unknown_Should_Throw()
        {
            // Arrange
            var factory = new AdapterFactory();
            var services = new ServiceCollection();

            // Act
            var action = () => factory.RegisterAdapter(services, PlatformType.Unknown);

            // Assert
            action.Should().Throw<InvalidOperationException>();
        }

        [Fact]
        public void GetRunningApplication_Should_Not_Throw()
        {
            // Arrange
            var factory = new AdapterFactory();

            // Act
            var action = () => factory.GetRunningApplication();

            // Assert
            action.Should().NotThrow();
        }
    }
}
