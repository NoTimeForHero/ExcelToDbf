using System;
using System.Collections.Generic;
using System.IO;
using System.Security;
using System.Windows;
using ExcelToDbf.Core.Models;
using ExcelToDbf.Utils;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using NLog;
using Screen = System.Windows.Forms.Screen;

namespace ExcelToDbf.Core.Services
{
    internal sealed class LastLaunchService
    {
        private readonly ILogger logger;
        private readonly string path = Constants.LastLaunchFile;
        private readonly LastLaunch session = new LastLaunch();

        public LastLaunchService(ILogger logger)
        {
            this.logger = logger;
        }

        public LastLaunch Load()
        {
            try
            {
                var settings = new JsonSerializerSettings { DateParseHandling = DateParseHandling.None };
                var document = JsonConvert.DeserializeObject<JObject>(File.ReadAllText(path), settings);
                if (document == null)
                {
                    logger.Warn($"Файл «{path}» не содержит данных последнего запуска.");
                    return session;
                }

                session.Path = ReadField<string>(document, nameof(LastLaunch.Path));
                session.Results = ReadField<List<ConvertService.Result>>(document, nameof(LastLaunch.Results));
                session.Results?.RemoveAll(result => result == null);
                session.WindowPlacement = ReadField<WindowPlacement>(document, nameof(LastLaunch.WindowPlacement));
            }
            catch (FileNotFoundException)
            {
                // A first launch has no saved state.
            }
            catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException ||
                                       ex is SecurityException || ex is JsonException)
            {
                logger.Warn(ex, $"Не удалось прочитать файл «{path}».");
            }

            return session;
        }

        public void TrackWindow(Window window, Func<string> getSelectedPath)
        {
            WindowPlacement closingPlacement = null;
            window.WindowStartupLocation = WindowStartupLocation.Manual;
            window.SourceInitialized += (sender, args) => RestoreWindow(window);
            window.Closing += (sender, args) => closingPlacement = CaptureWindow(window);
            window.Closed += (sender, args) =>
            {
                session.WindowPlacement = closingPlacement ?? session.WindowPlacement;
                var selectedPath = getSelectedPath();
                if (!string.IsNullOrEmpty(selectedPath))
                    session.Path = selectedPath;
                Save();
            };
        }

        public void SaveResults(string selectedPath, List<ConvertService.Result> results)
        {
            session.Path = selectedPath;
            session.Results = results;
            Save();
        }

        private T ReadField<T>(JObject document, string name) where T : class
        {
            try
            {
                return document[name]?.ToObject<T>();
            }
            catch (Exception ex) when (ex is JsonException || ex is ArgumentException ||
                                       ex is InvalidOperationException || ex is FormatException)
            {
                logger.Warn(ex, $"Не удалось прочитать поле «{name}» из файла «{path}».");
                return null;
            }
        }

        private void Save()
        {
            var temporaryPath = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
            try
            {
                FileStorage.Save(temporaryPath, session);
                if (File.Exists(path))
                    File.Replace(temporaryPath, path, null);
                else
                    File.Move(temporaryPath, path);
            }
            catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException ||
                                       ex is SecurityException || ex is JsonException)
            {
                logger.Warn(ex, $"Не удалось сохранить файл «{path}».");
                try
                {
                    File.Delete(temporaryPath);
                }
                catch (Exception cleanupError) when (cleanupError is IOException ||
                                                     cleanupError is UnauthorizedAccessException ||
                                                     cleanupError is SecurityException)
                {
                    logger.Warn(cleanupError, $"Не удалось удалить временный файл «{temporaryPath}».");
                }
            }
        }

        private void RestoreWindow(Window window)
        {
            var target = PresentationSource.FromVisual(window)?.CompositionTarget;
            var primary = Screen.PrimaryScreen;
            if (target == null || primary == null)
                return;

            var placement = session.WindowPlacement;
            Screen savedMonitor = null;
            var savedBounds = Rect.Empty;
            if (placement != null && IsValidBounds(placement.Left, placement.Top, placement.Width, placement.Height))
            {
                savedBounds = new Rect(placement.Left, placement.Top, placement.Width, placement.Height);
                savedMonitor = FindMonitor(savedBounds);
            }

            var area = (savedMonitor ?? primary).WorkingArea;
            if (area.Width <= 0 || area.Height <= 0)
                return;

            var workArea = new Rect(area.X, area.Y, area.Width, area.Height);
            var bounds = savedBounds.IsEmpty
                ? Rect.Transform(new Rect(0, 0, window.Width, window.Height), target.TransformToDevice)
                : savedBounds;
            var width = Math.Min(bounds.Width, workArea.Width);
            var height = Math.Min(bounds.Height, workArea.Height);
            var left = workArea.Left + (workArea.Width - width) / 2;
            var top = workArea.Top + (workArea.Height - height) / 2;

            if (savedMonitor != null)
            {
                left = Math.Max(workArea.Left, Math.Min(savedBounds.Left, workArea.Right - width));
                top = Math.Max(workArea.Top, Math.Min(savedBounds.Top, workArea.Bottom - height));
            }

            var restored = Rect.Transform(new Rect(left, top, width, height), target.TransformFromDevice);
            window.Width = restored.Width;
            window.Height = restored.Height;
            window.Left = restored.Left;
            window.Top = restored.Top;
        }

        private static Screen FindMonitor(Rect bounds)
        {
            Screen bestMonitor = null;
            double largestOverlap = 0;
            foreach (var screen in Screen.AllScreens)
            {
                var intersection = Rect.Intersect(bounds, GetMonitorBounds(screen));
                if (intersection.IsEmpty)
                    continue;

                var overlap = intersection.Width * intersection.Height;
                if (overlap > largestOverlap)
                {
                    largestOverlap = overlap;
                    bestMonitor = screen;
                }
            }

            return bestMonitor;
        }

        private static Rect GetMonitorBounds(Screen monitor)
        {
            var bounds = monitor.Bounds;
            return new Rect(bounds.X, bounds.Y, bounds.Width, bounds.Height);
        }

        private static WindowPlacement CaptureWindow(Window window)
        {
            var target = PresentationSource.FromVisual(window)?.CompositionTarget;
            var bounds = window.RestoreBounds;
            if (target == null || bounds.IsEmpty)
                return null;

            bounds.Transform(target.TransformToDevice);
            if (!IsValidBounds(bounds.Left, bounds.Top, bounds.Width, bounds.Height))
                return null;

            return new WindowPlacement
            {
                Left = bounds.Left,
                Top = bounds.Top,
                Width = bounds.Width,
                Height = bounds.Height
            };
        }

        private static bool IsValidBounds(double left, double top, double width, double height)
        {
            return IsFinite(left) && IsFinite(top) && IsFinite(width) && IsFinite(height) &&
                   width >= 1 && height >= 1 && IsFinite(left + width) && IsFinite(top + height);
        }

        private static bool IsFinite(double value)
        {
            return !double.IsNaN(value) && !double.IsInfinity(value);
        }
    }
}
