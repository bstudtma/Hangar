// <copyright file="SimConnectDataVelocityWorld.cs" company="BARS">
// Copyright (c) BARS. All rights reserved.
// </copyright>

using SimConnect.NET;
using SimConnect.NET.SimVar;

namespace Hangar
{
    /// <summary>
    /// Groups the three world velocity SimVars so they can be sent/retrieved together in one SimConnect transaction.
    /// Units are feet per second as per SimConnect documentation.
    /// </summary>
    public struct SimConnectDataVelocityWorld
    {
        /// <summary>
        /// Velocity in the world X-axis (feet per second).
        /// </summary>
        [SimConnect("VELOCITY WORLD X", "Feet per second")]
        public double X;

        /// <summary>
        /// Velocity in the world Y-axis (feet per second).
        /// </summary>
        [SimConnect("VELOCITY WORLD Y", "Feet per second")]
        public double Y;

        /// <summary>
        /// Velocity in the world Z-axis (feet per second).
        /// </summary>
        [SimConnect("VELOCITY WORLD Z", "Feet per second")]
        public double Z;
    }
}
