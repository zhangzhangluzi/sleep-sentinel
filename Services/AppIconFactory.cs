using System.Drawing.Drawing2D;

namespace SleepSentinel.Services;

public static class AppIconFactory
{
    public static Icon CreateAppIcon(int size = 64)
    {
        using var bitmap = new Bitmap(size, size);
        using var graphics = Graphics.FromImage(bitmap);
        graphics.SmoothingMode = SmoothingMode.AntiAlias;
        graphics.Clear(Color.Transparent);

        var backgroundRect = new RectangleF(2, 2, size - 4, size - 4);
        using var backgroundBrush = new LinearGradientBrush(
            new PointF(0, 0),
            new PointF(size, size),
            Color.FromArgb(22, 36, 58),
            Color.FromArgb(19, 93, 112));
        graphics.FillEllipse(backgroundBrush, backgroundRect);

        // Crescent moon marks the sleep intent.
        using var moonBrush = new SolidBrush(Color.FromArgb(247, 212, 113));
        graphics.FillEllipse(moonBrush, size * 0.18f, size * 0.16f, size * 0.34f, size * 0.34f);
        using var cutBrush = new SolidBrush(Color.FromArgb(19, 93, 112));
        graphics.FillEllipse(cutBrush, size * 0.28f, size * 0.12f, size * 0.30f, size * 0.30f);

        // Shield hints that wake-ups are being guarded.
        using var shieldPath = new GraphicsPath();
        shieldPath.AddPolygon(new[]
        {
            new PointF(size * 0.50f, size * 0.24f),
            new PointF(size * 0.72f, size * 0.32f),
            new PointF(size * 0.70f, size * 0.56f),
            new PointF(size * 0.50f, size * 0.78f),
            new PointF(size * 0.30f, size * 0.56f),
            new PointF(size * 0.28f, size * 0.32f)
        });
        using var shieldBrush = new SolidBrush(Color.FromArgb(232, 244, 246));
        graphics.FillPath(shieldBrush, shieldPath);

        using var accentPen = new Pen(Color.FromArgb(14, 86, 104), Math.Max(2f, size * 0.05f))
        {
            StartCap = LineCap.Round,
            EndCap = LineCap.Round
        };
        graphics.DrawLine(accentPen, size * 0.39f, size * 0.47f, size * 0.61f, size * 0.47f);
        graphics.DrawLine(accentPen, size * 0.41f, size * 0.57f, size * 0.58f, size * 0.57f);

        using var borderPen = new Pen(Color.FromArgb(255, 255, 255, 28), Math.Max(1.5f, size * 0.03f));
        graphics.DrawEllipse(borderPen, backgroundRect);

        var icon = Icon.FromHandle(bitmap.GetHicon());
        return (Icon)icon.Clone();
    }
}
