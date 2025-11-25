using System;

namespace PPA.Core.Exceptions
{
    /// <summary>
    /// PPA 基础异常类
    /// </summary>
    public class PPAException : Exception
    {
        /// <summary>
        /// 错误代码
        /// </summary>
        public string ErrorCode { get; }

        public PPAException(string message)
            : base(message)
        {
        }

        public PPAException(string message, string errorCode)
            : base(message)
        {
            ErrorCode = errorCode;
        }

        public PPAException(string message, Exception innerException)
            : base(message, innerException)
        {
        }

        public PPAException(string message, string errorCode, Exception innerException)
            : base(message, innerException)
        {
            ErrorCode = errorCode;
        }
    }

    /// <summary>
    /// 平台不支持异常
    /// </summary>
    public class PlatformNotSupportedException : PPAException
    {
        public PlatformNotSupportedException(string platform)
            : base($"不支持的平台: {platform}", "PLATFORM_NOT_SUPPORTED")
        {
        }
    }

    /// <summary>
    /// 功能不支持异常
    /// </summary>
    public class FeatureNotSupportedException : PPAException
    {
        public FeatureNotSupportedException(string feature)
            : base($"当前平台不支持此功能: {feature}", "FEATURE_NOT_SUPPORTED")
        {
        }
    }

    /// <summary>
    /// 选择无效异常
    /// </summary>
    public class InvalidSelectionException : PPAException
    {
        public InvalidSelectionException(string message)
            : base(message, "INVALID_SELECTION")
        {
        }
    }

    /// <summary>
    /// 操作失败异常
    /// </summary>
    public class OperationFailedException : PPAException
    {
        public OperationFailedException(string operation, Exception innerException)
            : base($"操作失败: {operation}", "OPERATION_FAILED", innerException)
        {
        }
    }
}
