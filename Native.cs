using System;
using System.Runtime.InteropServices;

namespace Hangar;

internal static class NativeSimConnect
{
    // Map a client event ID to a SimConnect named event (e.g., LANDING_LIGHTS_OFF)
    [DllImport("SimConnect.dll")]
    public static extern int SimConnect_MapClientEventToSimEvent(
        IntPtr hSimConnect,
        uint eventId,
        [MarshalAs(UnmanagedType.LPStr)] string eventName);
}
