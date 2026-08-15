#include "my_application.h"

#include <cstdlib>

int main(int argc, char** argv) {
  // Force the X11 backend. Native Wayland doesn't let apps set their own
  // window icon or some client-side decoration details at runtime the way
  // X11/XWayland does, which is why the title bar and icon looked better
  // under GDK_BACKEND=x11. Must be set before GTK initializes.
  setenv("GDK_BACKEND", "x11", 1);

  g_autoptr(MyApplication) app = my_application_new();
  return g_application_run(G_APPLICATION(app), argc, argv);
}