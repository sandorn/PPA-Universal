using PPA.Core.Abstraction;

namespace PPA.Core.Abstraction
{
    /// <summary>
    /// 抽象的 idMso 命令执行器，用于在不同平台上执行 Ribbon 命令（如 ObjectEffectShadowGallery）。
    /// </summary>
    public interface IIdMsoCommandExecutor
    {
        /// <summary>
        /// 尝试在给定的应用程序上下文中执行指定的 idMso 命令。
        /// </summary>
        /// <param name="appContext">当前应用程序上下文（用于区分 PowerPoint / WPS 等平台）。</param>
        /// <param name="idMso">要执行的命令 ID，例如 "ObjectEffectShadowGallery"。</param>
        /// <returns>执行成功返回 true，失败或不支持时返回 false。</returns>
        bool TryExecute(IApplicationContext appContext, string idMso);
    }
}
