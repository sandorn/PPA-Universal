using System.Collections.Generic;
using PPA.Business.Abstractions;
using PPA.Business.Geometry;
using PPA.Core.Abstraction;
using Xunit;

namespace PPA.Business.Tests.Geometry
{
	public sealed class ShapeAlignmentMathTests
	{
		private static ShapeRect R(float left, float top, float w, float h) => new ShapeRect(left, top, w, h);

		[Fact]
		public void CalculateSelectionReference_Left_is_min_left()
		{
			var shapes = new List<ShapeRect> { R(10, 0, 5, 5), R(3, 0, 2, 2), R(20, 0, 1, 1) };
			Assert.Equal(3f, ShapeAlignmentMath.CalculateSelectionReference(shapes, AlignmentType.Left));
		}

		[Fact]
		public void CalculateSelectionReference_CenterHorizontal_is_mid_of_span()
		{
			var shapes = new List<ShapeRect> { R(0, 0, 10, 1), R(20, 0, 10, 1) };
			Assert.Equal(15f, ShapeAlignmentMath.CalculateSelectionReference(shapes, AlignmentType.CenterHorizontal));
		}

		[Fact]
		public void CalculateSlideReference_Right_uses_slide_width()
		{
			var fallback = new List<ShapeRect> { R(0, 0, 1, 1) };
			Assert.Equal(960f, ShapeAlignmentMath.CalculateSlideReference(AlignmentType.Right, 960f, 540f, fallback));
		}

		[Fact]
		public void ApplyAlignment_Left_sets_left_to_ref()
		{
			var b = R(50, 10, 20, 30);
			var r = ShapeAlignmentMath.ApplyAlignment(b, AlignmentType.Left, 100f);
			Assert.Equal(100f, r.Left);
			Assert.Equal(10f, r.Top);
			Assert.Equal(20f, r.Width);
			Assert.Equal(30f, r.Height);
		}

		[Fact]
		public void ApplyAlignment_Right_keeps_size()
		{
			var b = R(50, 0, 20, 10);
			var r = ShapeAlignmentMath.ApplyAlignment(b, AlignmentType.Right, 100f);
			Assert.Equal(80f, r.Left);
			Assert.Equal(100f, r.Right);
		}

		[Fact]
		public void SnapOtherToReference_Right_places_other_left_at_reference_right()
		{
			var reference = R(0, 0, 10, 10);
			var other = R(100, 5, 8, 6);
			var r = ShapeAlignmentMath.SnapOtherToReference(other, reference, SnapDirection.Right);
			Assert.Equal(10f, r.Left);
			Assert.Equal(18f, r.Right);
			Assert.Equal(5f, r.Top);
		}

		[Fact]
		public void SwapCenters_exchanges_centers_preserving_sizes()
		{
			var a = R(0, 0, 10, 10);
			var b = R(50, 20, 20, 10);
			var (na, nb) = ShapeAlignmentMath.SwapCenters(a, b);
			Assert.Equal(60f, na.CenterX);
			Assert.Equal(25f, na.CenterY);
			Assert.Equal(10f, na.Width);
			Assert.Equal(10f, na.Height);
			Assert.Equal(5f, nb.CenterX);
			Assert.Equal(5f, nb.CenterY);
			Assert.Equal(20f, nb.Width);
			Assert.Equal(10f, nb.Height);
		}

		[Fact]
		public void ComputeHorizontalDistributedLefts_three_rects_equal_gap()
		{
			var sorted = new List<ShapeRect>
			{
				R(0, 0, 10, 5),
				R(20, 0, 10, 5),
				R(40, 0, 10, 5)
			};
			var lefts = ShapeAlignmentMath.ComputeHorizontalDistributedLefts(sorted);
			Assert.Equal(3, lefts.Count);
			Assert.Equal(0f, lefts[0]);
			Assert.Equal(20f, lefts[1]);
			Assert.Equal(40f, lefts[2]);
		}

		[Fact]
		public void ComputeVerticalDistributedTops_redistributes_gap()
		{
			var sorted = new List<ShapeRect>
			{
				R(0, 0, 5, 10),
				R(0, 15, 5, 10),
				R(0, 50, 5, 10)
			};
			var tops = ShapeAlignmentMath.ComputeVerticalDistributedTops(sorted);
			Assert.Equal(0f, tops[0]);
			Assert.Equal(25f, tops[1]);
			Assert.Equal(50f, tops[2]);
		}
	}
}
