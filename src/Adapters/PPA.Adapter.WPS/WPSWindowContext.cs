using PPA.Core.Abstraction;

namespace PPA.Adapter.WPS
{
    /// <summary>
    /// WPS 窗口上下文实现
    /// </summary>
    public class WPSWindowContext : IWindowContext
    {
        private readonly dynamic _window;

        public WPSWindowContext(dynamic window)
        {
            _window = window;
        }

        public ViewType ViewType
        {
            get
            {
                try
                {
                    int viewType = _window?.ViewType ?? 1;
                    switch (viewType)
                    {
                        case 1: // ppViewNormal
                        case 9:
                            return ViewType.Normal;
                        case 6: // ppViewOutline
                            return ViewType.Outline;
                        case 4: // ppViewSlideSorter
                            return ViewType.SlideSorter;
                        case 5: // ppViewNotesPage
                            return ViewType.NotesPage;
                        case 7: // ppViewTitleMaster
                        case 2: // ppViewSlideMaster
                        case 3: // ppViewHandoutMaster
                        case 8: // ppViewNotesMaster
                            return ViewType.Master;
                        default:
                            return ViewType.Normal;
                    }
                }
                catch
                {
                    return ViewType.Normal;
                }
            }
        }

        public ISlideContext ActiveSlide
        {
            get
            {
                try
                {
                    dynamic slide = _window?.View?.Slide;
                    return slide != null ? new WPSSlideContext(slide) : null;
                }
                catch
                {
                    return null;
                }
            }
        }

        public int Zoom
        {
            get
            {
                try { return _window?.View?.Zoom ?? 100; }
                catch { return 100; }
            }
            set
            {
                try { if (_window?.View != null) _window.View.Zoom = value; }
                catch { }
            }
        }

        public object NativeWindow => _window;
    }
}
