// -----------------------------------------------------------------------------
// Ported from packages/dds/merge-tree/src/constants.ts
// -----------------------------------------------------------------------------

#nullable enable

namespace Microsoft.Office.Web.Fluid.MergeTree
{
    /// <summary>
    /// Foundational constants used by merge-tree sequencing logic.
    /// </summary>
    public static class Constants
    {
        /// <summary>
        /// The sequence number which can be seen by all ops.
        /// </summary>
        public const long UniversalSequenceNumber = 0;

        /// <summary>
        /// The sequence number of an op before it is acked.
        /// </summary>
        public const long UnassignedSequenceNumber = -1;

        /// <summary>
        /// The sequence number used by tree maintenance operations.
        /// </summary>
        public const long TreeMaintenanceSequenceNumber = -2;

        /// <summary>
        /// The client id for local operations.
        /// </summary>
        public const int LocalClientId = -1;

        /// <summary>
        /// The client id for non-collaborative operations.
        /// </summary>
        public const int NonCollabClient = -2;

        /// <summary>
        /// Used as the client id for operations that were squashed upon resubmission and should therefore never be seen by other clients.
        /// </summary>
        public const int SquashClient = -3;
    }
}
