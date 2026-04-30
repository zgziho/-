using System;
using System.IO;
using System.Linq;
using System.Reflection;

namespace WpfApp1.Service
{
    public static class NativeLibraryLoader
    {
        private static bool _initialized;

        public static void Initialize()
        {
            if (_initialized)
                return;

            var assembly = Assembly.GetExecutingAssembly();
            var resourceName = assembly.GetManifestResourceNames()
                .FirstOrDefault(name => name.EndsWith("zlgcan.dll", StringComparison.OrdinalIgnoreCase));
            
            if (string.IsNullOrEmpty(resourceName))
            {
                var availableResources = string.Join(", ", assembly.GetManifestResourceNames());
                throw new FileNotFoundException($"嵌入资源 zlgcan.dll 未找到。可用资源: {availableResources}");
            }
            
            using var stream = assembly.GetManifestResourceStream(resourceName);
            if (stream == null)
                throw new FileNotFoundException("无法打开嵌入资源流", resourceName);

            var appBaseDir = AppDomain.CurrentDomain.BaseDirectory;
            var dllPath = Path.Combine(appBaseDir, "zlgcan.dll");
            
            using var fileStream = File.Create(dllPath);
            stream.CopyTo(fileStream);

            _initialized = true;
        }
    }
}