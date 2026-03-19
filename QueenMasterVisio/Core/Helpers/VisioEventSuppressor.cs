using System;

namespace QueenMasterVisio.Core.Helpers
{
    /// <summary>
    /// Централизованный супрессор событий Visio.
    /// Позволяет временно отключать onShapeAdded, onShapeChanged и т.д.
    /// </summary>
    public static class VisioEventSuppressor
    {
        private static int _shapeAddedCount = 0;

        /// <summary>Проверяет, сейчас ли мы в режиме "бот работает"</summary>
        public static bool IsShapeAddedSuppressed => _shapeAddedCount > 0;

        /// <summary>
        /// Используй так:
        /// using (VisioEventSuppressor.SuppressShapeAdded())
        /// {
        ///     page.Paste(...);
        ///     // любой код, который добавляет/меняет фигуры
        /// }
        /// </summary>
        public static IDisposable SuppressShapeAdded() => new ShapeAddedSuppressor();

        // Внутренний класс-супрессор
        private sealed class ShapeAddedSuppressor : IDisposable
        {
            short oldEvents = VisioEventAggregator.app.EventsEnabled;
            public ShapeAddedSuppressor()
            {
                _shapeAddedCount++;
                VisioEventAggregator.app.Application.EventsEnabled = 0;
            }

            public void Dispose()
            {
                _shapeAddedCount--;
                VisioEventAggregator.app.EventsEnabled = oldEvents;
            }
        }
    }
}