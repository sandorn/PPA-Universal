using System.Collections.Generic;
using System.Linq;
using PPA.Business.Abstractions;
using PPA.Core.Abstraction;

namespace PPA.Business.Geometry
{
	/// <summary>
	/// 对齐 / 吸附 / 分布等纯几何计算（无 COM、无日志），供 <see cref="Services.AlignmentService"/> 与单元测试复用。
	/// </summary>
	public static class ShapeAlignmentMath
	{
		public static float CalculateSelectionReference(IReadOnlyList<ShapeRect> shapes, AlignmentType alignment)
		{
			if (shapes == null || shapes.Count == 0)
				return 0f;

			switch (alignment)
			{
				case AlignmentType.Left:
					return shapes.Min(s => s.Left);
				case AlignmentType.Right:
					return shapes.Max(s => s.Right);
				case AlignmentType.Top:
					return shapes.Min(s => s.Top);
				case AlignmentType.Bottom:
					return shapes.Max(s => s.Bottom);
				case AlignmentType.CenterHorizontal:
					{
						var minX = shapes.Min(s => s.Left);
						var maxX = shapes.Max(s => s.Right);
						return (minX + maxX) / 2;
					}
				case AlignmentType.CenterVertical:
					{
						var minY = shapes.Min(s => s.Top);
						var maxY = shapes.Max(s => s.Bottom);
						return (minY + maxY) / 2;
					}
				default:
					return 0f;
			}
		}

		public static float CalculateSlideReference(
			AlignmentType alignment,
			float slideWidth,
			float slideHeight,
			IReadOnlyList<ShapeRect> shapesFallback)
		{
			switch (alignment)
			{
				case AlignmentType.Left:
				case AlignmentType.Top:
					return 0f;
				case AlignmentType.Right:
					return slideWidth;
				case AlignmentType.Bottom:
					return slideHeight;
				case AlignmentType.CenterHorizontal:
					return slideWidth / 2;
				case AlignmentType.CenterVertical:
					return slideHeight / 2;
				default:
					return CalculateSelectionReference(shapesFallback, alignment);
			}
		}

		public static float CalculateSingleShapeReference(ShapeRect bounds, AlignmentType alignment)
		{
			switch (alignment)
			{
				case AlignmentType.Left:
					return bounds.Left;
				case AlignmentType.Right:
					return bounds.Right;
				case AlignmentType.Top:
					return bounds.Top;
				case AlignmentType.Bottom:
					return bounds.Bottom;
				case AlignmentType.CenterHorizontal:
					return bounds.CenterX;
				case AlignmentType.CenterVertical:
					return bounds.CenterY;
				default:
					return 0f;
			}
		}

		public static ShapeRect ApplyAlignment(ShapeRect bounds, AlignmentType alignment, float refValue)
		{
			switch (alignment)
			{
				case AlignmentType.Left:
					return new ShapeRect(refValue, bounds.Top, bounds.Width, bounds.Height);
				case AlignmentType.Right:
					return new ShapeRect(refValue - bounds.Width, bounds.Top, bounds.Width, bounds.Height);
				case AlignmentType.Top:
					return new ShapeRect(bounds.Left, refValue, bounds.Width, bounds.Height);
				case AlignmentType.Bottom:
					return new ShapeRect(bounds.Left, refValue - bounds.Height, bounds.Width, bounds.Height);
				case AlignmentType.CenterHorizontal:
					return new ShapeRect(refValue - bounds.Width / 2, bounds.Top, bounds.Width, bounds.Height);
				case AlignmentType.CenterVertical:
					return new ShapeRect(bounds.Left, refValue - bounds.Height / 2, bounds.Width, bounds.Height);
				default:
					return bounds;
			}
		}

		/// <summary>将 <paramref name="other"/> 贴靠到 <paramref name="reference"/> 的指定侧（与 <see cref="IAlignmentService.SnapToShape"/> 语义一致）。</summary>
		public static ShapeRect SnapOtherToReference(ShapeRect other, ShapeRect reference, SnapDirection snapDirection)
		{
			var dx = other.Right - other.Left;
			var dy = other.Bottom - other.Top;

			return snapDirection switch
			{
				SnapDirection.Left => new ShapeRect(reference.Left - dx, other.Top, dx, dy),
				SnapDirection.Right => new ShapeRect(reference.Right, other.Top, dx, dy),
				SnapDirection.Top => new ShapeRect(other.Left, reference.Top - dy, dx, dy),
				SnapDirection.Bottom => new ShapeRect(other.Left, reference.Bottom, dx, dy),
				_ => other
			};
		}

		/// <summary>交换两矩形中心位置，各自宽高不变。</summary>
		public static (ShapeRect first, ShapeRect second) SwapCenters(ShapeRect a, ShapeRect b)
		{
			var c1x = a.CenterX;
			var c1y = a.CenterY;
			var c2x = b.CenterX;
			var c2y = b.CenterY;

			var na = new ShapeRect(
				c2x - a.Width / 2,
				c2y - a.Height / 2,
				a.Width,
				a.Height);
			var nb = new ShapeRect(
				c1x - b.Width / 2,
				c1y - b.Height / 2,
				b.Width,
				b.Height);
			return (na, nb);
		}

		/// <summary>水平均匀分布：按 Left 排序后的形状，返回与 <paramref name="sortedByLeft"/> 同序的新 Left。</summary>
		public static IReadOnlyList<float> ComputeHorizontalDistributedLefts(IReadOnlyList<ShapeRect> sortedByLeft)
		{
			if (sortedByLeft == null || sortedByLeft.Count < 2)
				return sortedByLeft?.Select(r => r.Left).ToList() ?? new List<float>();

			var leftmost = sortedByLeft[0].Left;
			var rightmost = sortedByLeft[sortedByLeft.Count - 1].Right;
			var totalShapeWidth = sortedByLeft.Sum(s => s.Width);
			var totalSpace = rightmost - leftmost - totalShapeWidth;
			var gap = totalSpace / (sortedByLeft.Count - 1);

			var result = new List<float>(sortedByLeft.Count);
			float currentX = leftmost;
			foreach (var r in sortedByLeft)
			{
				result.Add(currentX);
				currentX += r.Width + gap;
			}

			return result;
		}

		/// <summary>垂直均匀分布：按 Top 排序后的形状，返回与排序列表同序的新 Top。</summary>
		public static IReadOnlyList<float> ComputeVerticalDistributedTops(IReadOnlyList<ShapeRect> sortedByTop)
		{
			if (sortedByTop == null || sortedByTop.Count < 2)
				return sortedByTop?.Select(r => r.Top).ToList() ?? new List<float>();

			var topmost = sortedByTop[0].Top;
			var bottommost = sortedByTop[sortedByTop.Count - 1].Bottom;
			var totalShapeHeight = sortedByTop.Sum(s => s.Height);
			var totalSpace = bottommost - topmost - totalShapeHeight;
			var gap = totalSpace / (sortedByTop.Count - 1);

			var result = new List<float>(sortedByTop.Count);
			float currentY = topmost;
			foreach (var r in sortedByTop)
			{
				result.Add(currentY);
				currentY += r.Height + gap;
			}

			return result;
		}
	}
}
