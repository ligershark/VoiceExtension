// Guids.cs
// MUST match guids.h
using System;

namespace MadsKristensen.VoiceExtension
{
    static class GuidList
    {
        public const string guidVoiceExtensionPkgString = "b4558cd7-da41-47e7-8969-46c357a1b8b3";
        public const string guidVoiceExtensionCmdSetString = "9ae4e318-bbcc-41c1-a880-e64ac747659b";

        public static readonly Guid guidVoiceExtensionCmdSet = new Guid(guidVoiceExtensionCmdSetString);
    };
}