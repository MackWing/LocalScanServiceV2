namespace LocalScanServiceV2.WIA
{
    public static class WiaConstants
    {
        // WIA 属性 ID
        public const int WIA_IPA_DATATYPE = 6144;
        public const int WIA_IPA_DEPTH = 6145;
        public const int WIA_IPA_COLOR_MODE = 6146;
        public const int WIA_IPS_XRES = 6147;
        public const int WIA_IPS_YRES = 6148;
        public const int WIA_IPS_XPOS = 6149;
        public const int WIA_IPS_YPOS = 6150;
        public const int WIA_IPS_XEXTENT = 6151;
        public const int WIA_IPS_YEXTENT = 6152;
        public const int WIA_IPS_BRIGHTNESS = 6154;
        public const int WIA_IPS_CONTRAST = 6155;
        public const int WIA_DPS_DOCUMENT_HANDLING_STATUS = 3086;
        public const int WIA_DPS_DOCUMENT_HANDLING_CAPABILITIES = 3087;
        public const int WIA_DPS_DOCUMENT_HANDLING_SELECT = 3088;
        public const int WIA_DPS_PAGE_SIZE = 3089;
        public const int WIA_DPS_PAGE_WIDTH = 3090;
        public const int WIA_DPS_PAGE_HEIGHT = 3091;
        public const int WIA_DPS_SCAN_AHEAD_PAGES = 3092;
        public const int WIA_DPS_MAX_SCAN_TIME = 3093;
        public const int WIA_DPS_PAGES = 3094;
        public const int WIA_DPS_PAGES_PER_FEED = 3095;
        public const int WIA_DPS_PAGES_EXIST = 3096;
        public const int WIA_DPS_AUTO_FEED = 3097;
        public const int WIA_DPS_CURRENT_PAGE = 3098;
        public const int WIA_DPS_TOTAL_PAGES = 3099;
        public const int WIA_DPS_FEED_READY = 3100;

        // 颜色模式值
        public const int WIA_COLOR_MODE_COLOR = 1;
        public const int WIA_COLOR_MODE_GRAYSCALE = 2;
        public const int WIA_COLOR_MODE_BLACKANDWHITE = 4;

        // 纸张来源值
        public const int WIA_PAPER_SOURCE_FLATBED = 1;
        public const int WIA_PAPER_SOURCE_FEEDER = 256;

        // 设备类型
        public const int WIA_DEVICE_SCANNER = 6;

        // 错误码
        public const uint WIA_ERROR_BUSY = 0x80210006;
        public const uint WIA_ERROR_DEVICE_COMMUNICATION = 0x80210004;
        public const uint WIA_ERROR_OFFLINE = 0x80210003;
        public const uint WIA_ERROR_PAPER_JAM = 0x80210005;
        public const uint WIA_ERROR_PAPER_EMPTY = 0x80210002;
        public const uint WIA_ERROR_NO_DEVICE_AVAILABLE = 0x80210001;
    }
}