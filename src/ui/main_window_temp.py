 def on_window_resize(self, event=None):
      """
        Handle window resize events for responsive UI

        Args:
            event: The window resize event
        """
       # Only respond to root window resizes, not child widget resizes
       if event and event.widget != self.root:
            return

        # Update any responsive UI elements
        # For example, adjust the wraplength of the details label based on window size
        if hasattr(self, 'details_label'):
            # Make the wraplength responsive to window width
            window_width = self.root.winfo_width()
            # Use 70% of the window width for the details label wraplength
            self.details_label.configure(wraplength=int(window_width * 0.7))

        # You can add more responsive adjustments here
