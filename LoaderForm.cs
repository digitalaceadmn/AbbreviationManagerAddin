using System;
using System.Drawing;
using System.Windows.Forms;

public class LoaderForm : Form
{
    private Label lblTitle;
    private Label lblMessage;
    private ProgressBar progressBar;
    private Panel container;

    public LoaderForm()
    {
        this.FormBorderStyle = FormBorderStyle.None;
        this.StartPosition = FormStartPosition.CenterScreen;
        this.Width = 420;
        this.Height = 160;
        this.TopMost = true;
        this.BackColor = Color.FromArgb(240, 240, 240);

        container = new Panel
        {
            Dock = DockStyle.Fill,
            BackColor = Color.White,
            Padding = new Padding(20)
        };
        this.Controls.Add(container);

        lblTitle = new Label
        {
            Text = "Processing document",
            Dock = DockStyle.Top,
            Height = 28,
            Font = new Font("Segoe UI", 11F, FontStyle.Bold),
            ForeColor = Color.FromArgb(50, 50, 50),
            TextAlign = ContentAlignment.MiddleLeft
        };

        lblMessage = new Label
        {
            Text = "Replacing abbreviations. Please wait…",
            Dock = DockStyle.Top,
            Height = 30,
            Font = new Font("Segoe UI", 9.5F),
            ForeColor = Color.FromArgb(90, 90, 90),
            Padding = new Padding(0, 6, 0, 0)
        };

        progressBar = new ProgressBar
        {
            Style = ProgressBarStyle.Marquee,
            Dock = DockStyle.Bottom,
            Height = 18,
            MarqueeAnimationSpeed = 28
        };

        container.Controls.Add(progressBar);
        container.Controls.Add(lblMessage);
        container.Controls.Add(lblTitle);

        this.Paint += (s, e) =>
        {
            using (Pen p = new Pen(Color.FromArgb(220, 220, 220)))
            {
                e.Graphics.DrawRectangle(p, 0, 0, Width - 1, Height - 1);
            }
        };
    }
}
