// Guids.cs
// MUST match guids.h

using System;

namespace MidnightDevelopers.VisualStudio.VsRestart
{
    static class GuidList
    {
        public const string VsRestarterPackageId = "7c8887d9-d300-4e91-aadb-862a19a081ff";
        public const string RestartElevatedCommandGroupId = "15dc28d5-04f4-4698-90e0-e3e16bc6894f";

        public const string TopLevelMenuGroupId = "D2FB6644-0147-4FDB-8F35-22B5F0AA8594";

        public static readonly Guid RestartElevatedGroupGuid = new Guid(RestartElevatedCommandGroupId);
        public static readonly Guid TopLevelMenuGuid = new Guid(TopLevelMenuGroupId);
    }
}