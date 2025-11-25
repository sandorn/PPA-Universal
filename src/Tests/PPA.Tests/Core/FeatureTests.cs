using FluentAssertions;
using PPA.Core.Abstraction;
using Xunit;

namespace PPA.Tests.Core
{
    /// <summary>
    /// Feature 枚举测试
    /// </summary>
    public class FeatureTests
    {
        [Theory]
        [InlineData(Feature.TableBasic)]
        [InlineData(Feature.TableAdvancedBorder)]
        [InlineData(Feature.ShapeAlignment)]
        [InlineData(Feature.ShapeBatch)]
        [InlineData(Feature.Chart)]
        [InlineData(Feature.TextAdvanced)]
        public void Feature_Values_Should_Be_Defined(Feature feature)
        {
            // Assert
            feature.Should().BeDefined();
        }
    }
}
