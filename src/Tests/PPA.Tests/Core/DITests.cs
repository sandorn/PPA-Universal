using FluentAssertions;
using Microsoft.Extensions.DependencyInjection;
using PPA.Adapter.PowerPoint.DI;
using PPA.Business.Abstractions;
using PPA.Business.DI;
using PPA.Core.DI;
using PPA.Logging;
using Xunit;

namespace PPA.Tests.Core
{
    /// <summary>
    /// 依赖注入测试
    /// </summary>
    public class DITests
    {
        [Fact]
        public void AddPPACore_Should_Register_Logger()
        {
            // Arrange
            var services = new ServiceCollection();

            // Act
            services.AddPPACore();
            var provider = services.BuildServiceProvider();

            // Assert
            var logger = provider.GetService<ILogger>();
            logger.Should().NotBeNull();
        }

        [Fact]
        public void AddPPABusiness_Should_Register_Services()
        {
            // Arrange
            var services = new ServiceCollection();

            // Act
            services.AddPPACore();
            services.AddPowerPointAdapter(); // Business 服务依赖适配器
            services.AddPPABusiness();
            var provider = services.BuildServiceProvider();

            // Assert
            var tableService = provider.GetService<ITableFormatService>();
            var alignService = provider.GetService<IAlignmentService>();

            tableService.Should().NotBeNull();
            alignService.Should().NotBeNull();
        }

        [Fact]
        public void Services_Should_Be_Resolvable_Multiple_Times()
        {
            // Arrange
            var services = new ServiceCollection();
            services.AddPPACore();
            services.AddPowerPointAdapter(); // Business 服务依赖适配器
            services.AddPPABusiness();
            var provider = services.BuildServiceProvider();

            // Act
            var service1 = provider.GetService<ITableFormatService>();
            var service2 = provider.GetService<ITableFormatService>();

            // Assert
            service1.Should().NotBeNull();
            service2.Should().NotBeNull();
        }
    }
}
