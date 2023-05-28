﻿using System;
using System.ComponentModel;

namespace ConformU
{
    /// <summary>
    /// Value to be used in the command line "Full Test" mode.
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public class FullTestValueAttribute : Attribute
    {
        /// <summary>
        /// Constructor that accepts and saves the full test value
        /// </summary>
        /// <param name="fullSettingsValue"></param>
        public FullTestValueAttribute(bool fullSettingsValue)
        {
            FullSettingsValue = fullSettingsValue;
        }

        /// <summary>
        /// Value to be used in command line "Full Settings" mode.
        /// </summary>
        public bool FullSettingsValue { get; set; }
    }
}