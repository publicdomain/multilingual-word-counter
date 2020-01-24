// <copyright file="MessageForm.cs" company="PublicDomain.com">
//     CC0 1.0 Universal (CC0 1.0) - Public Domain Dedication
//     https://creativecommons.org/publicdomain/zero/1.0/legalcode
// </copyright>

namespace MultilingualWordCounter
{
    // Directives
    using System;
    using System.Drawing;
    using System.Runtime.InteropServices;
    using System.Text.RegularExpressions;
    using System.Windows.Forms;

    /// <summary>
    /// Description of MessageForm.
    /// </summary>
    public partial class MessageForm : Form
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="T:MultilingualWordCounter.MessageForm"/> class.
        /// </summary>
        public MessageForm()
        {
            // The InitializeComponent() call is required for Windows Forms designer support.
            this.InitializeComponent();
        }

        /// <summary>
        /// Shows the message.
        /// </summary>
        /// <param name="message">The text.</param>
        /// <param name="title">The title.</param>
        public void ShowMessage(string message, string title)
        {
            // Set title
            this.Text = title;

            // Set message
            this.messageLabel.Text = message;

            // Show form topmost
            this.Show();
        }

        /// <summary>
        /// Handles the message form form closing event.
        /// </summary>
        /// <param name="sender">Sender object.</param>
        /// <param name="e">Event arguments.</param>
        private void OnMessageFormFormClosing(object sender, FormClosingEventArgs e)
        {
            // Hide form
            this.Hide();

            // Prevent closing
            e.Cancel = true;
        }

        /// <summary>
        /// Handles the close button click event.
        /// </summary>
        /// <param name="sender">Sender object.</param>
        /// <param name="e">Event arguments.</param>
        private void OnCloseButtonClick(object sender, EventArgs e)
        {
            // Hide via close
            this.Close();
        }
    }
}
