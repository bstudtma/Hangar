// <copyright file="SimConnectDataInitPosition.cs" company="BARS">
// Copyright (c) BARS. All rights reserved.
// </copyright>

using SimConnect.NET;
using SimConnect.NET.SimVar;

namespace Hangar
{
    /// <summary>
    /// The SimConnectDataInitPosition struct is used to initialize the position of the user aircraft, AI-controlled aircraft, or other simulation object.
    /// </summary>
    public struct SimConnectDataInitPosition
    {
        /// <summary>
        /// Gets or sets the latitude in degrees.
        /// </summary>
        [SimConnect("PLANE LATITUDE", "degrees")]
        public double Latitude;

        /// <summary>
        /// Gets or sets the longitude in degrees.
        /// </summary>
        [SimConnect("PLANE LONGITUDE", "degrees")]
        public double Longitude;

        /// <summary>
        /// Gets or sets the altitude in feet.
        /// </summary>
        [SimConnect("PLANE ALTITUDE", "feet")]
        public double Altitude;

        /// <summary>
        /// Gets or sets the pitch in degrees.
        /// </summary>
        [SimConnect("PLANE PITCH DEGREES", "degrees")]
        public double Pitch;

        /// <summary>
        /// Gets or sets the bank in degrees.
        /// </summary>
        [SimConnect("PLANE BANK DEGREES", "degrees")]
        public double Bank;

        /// <summary>
        /// Gets or sets the heading in degrees.
        /// </summary>
        [SimConnect("PLANE HEADING DEGREES TRUE", "degrees")]
        public double Heading;

        /// <summary>
        /// Gets or sets a value indicating whether the object is on the ground (1) or airborne (0).
        /// </summary>
        [SimConnect("SIM ON GROUND", "Bool")]
        public uint OnGround;

        /// <summary>
        /// Gets or sets the airspeed in knots, or one of the following special values:
        /// - INITPOSITION_AIRSPEED_CRUISE (-1): The aircraft's design cruising speed.
        /// - INITPOSITION_AIRSPEED_KEEP (-2): Maintain the current airspeed.
        /// </summary>
        [SimConnect("AIRSPEED TRUE", "knots")]
        public uint Airspeed;
    }

    public struct SimConnectDataLatLonAltNew
    {
        /// <summary>
        /// Gets or sets the latitude in degrees.
        /// </summary>
        [SimConnect("PLANE LATITUDE", "degrees")]
        public double Latitude;

        /// <summary>
        /// Gets or sets the longitude in degrees.
        /// </summary>
        [SimConnect("PLANE LONGITUDE", "degrees")]
        public double Longitude;

        /// <summary>
        /// Gets or sets the altitude in feet.
        /// </summary>
        [SimConnect("PLANE ALTITUDE", "feet")]
        public double Altitude;
    }
}
