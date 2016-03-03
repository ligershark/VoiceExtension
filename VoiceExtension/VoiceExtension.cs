namespace MadsKristensen.VoiceExtension
{
    using System;
    
    /// <summary>
    /// Helper class that exposes all GUIDs used across VS Package.
    /// </summary>
    internal sealed partial class PackageGuids
    {
        public const string guidVoiceExtensionPkgString = "b4558cd7-da41-47e7-8969-46c357a1b8b3";
        public const string guidVoiceExtensionCmdSetString = "9ae4e318-bbcc-41c1-a880-e64ac747659b";
        public const string guidImagesString = "7f0cf8f2-9680-4746-ac3d-647d4019dd85";
        public static Guid guidVoiceExtensionPkg = new Guid(guidVoiceExtensionPkgString);
        public static Guid guidVoiceExtensionCmdSet = new Guid(guidVoiceExtensionCmdSetString);
        public static Guid guidImages = new Guid(guidImagesString);
    }
    /// <summary>
    /// Helper class that encapsulates all CommandIDs uses across VS Package.
    /// </summary>
    internal sealed partial class PackageIds
    {
        public const int MyMenuGroup = 0x1020;
        public const int cmdidMyCommand = 0x0100;
        public const int icon = 0x0001;
    }
}
