// -----------------------------------------------------------------------------
// SHIM — deleted on transfer to host repo.
//
// These interfaces are copied from the host repo file:
//   waccobalt/src/server/dss/DocumentSessionService.Core/Fluid/FluidDispatcher.cs
//
// On transfer, the host's real definitions take over and this file is discarded.
// The purpose here is only to let cs-out/*.cs compile in isolation.
//
// Visibility: shim fields/members are public (vs internal in host). Reason:
// consumers span multiple assemblies. See ../README.md.
// -----------------------------------------------------------------------------

using System;

namespace Microsoft.Office.Web.Fluid
{
    /// <summary>
    /// wire protocol uses low case names; keeping them for consistency
    /// and to avoid conflict between SequenceNumber class name and sequenceNumber
    /// field name
    /// </summary>
    public struct SequenceNumber
    {
        public long clientSequenceNumber;
        public long referenceSequenceNumber;
        public long sequenceNumber;

        public bool HasClientSequenceNumber => clientSequenceNumber != 0;
        public bool HasSequenceNumber => sequenceNumber != 0;

        /// <summary>Test-only factory. Host repo does not need this.</summary>
        public static SequenceNumber ForTesting(long clientSeq, long refSeq = 0, long seq = 0)
        {
            return new SequenceNumber
            {
                clientSequenceNumber = clientSeq,
                referenceSequenceNumber = refSeq,
                sequenceNumber = seq,
            };
        }
    }

    public enum OpOrigin
    {
        Local,
        Remote,
    }

    /// <summary>
    /// subset of info from SequencedDocumentMessage
    /// </summary>
    public class SequencedDocumentMessageDescriptor
    {
        public readonly OpOrigin Origin;
        public readonly SequenceNumber Seq;
        public readonly string? ClientId;

        public bool HasClientId => ClientId != null;

        public long RefSeq => Seq.referenceSequenceNumber;

        /// <summary>
        /// Test-friendly constructor (shim only). The host repo has richer
        /// constructors that build a descriptor from a real SequencedDocumentMessage.
        /// </summary>
        public SequencedDocumentMessageDescriptor(SequenceNumber seq, OpOrigin origin, string? clientId = null)
        {
            Origin = origin;
            ClientId = clientId;
            Seq = seq;
        }
    }

    /// <summary>
    /// Required interface for all DSS objects. Exposes ID for an object which is used an as address.
    ///
    /// Note on threading. Even so the server is multi-threaded, handling things with locks on multiple levels
    /// is too complicated. Instead we are going to serialize access to model on top level and only allow one
    /// operation at time. Call running outside top level, should only use immutable objects (TBD)
    /// </summary>
    public interface IFluidDataObject
    {
        string Id { get; }
    }

    /// <summary>
    /// handles messages for specific instance of an object
    /// </summary>
    public interface IFluidDataObjectMessageHandler
    {
        /// <summary>
        /// Called when a remote op arrives for this DDS. `opJson` is the serialized op.
        /// </summary>
        /// <remarks>
        /// DEMO SIGNATURE: JSON directly. Host repo passes a Bond-wrapped payload;
        /// reconcile on transfer (waccobalt owns the wire format).
        /// </remarks>
        void ProcessDataObjectOp(SequencedDocumentMessageDescriptor descriptor, string opJson);

        /// <summary>
        /// Called for attach messages. The server-side client typically doesn't act on
        /// attach; keep this a no-op unless you need to reject unexpected attaches.
        /// </summary>
        void ProcessDataObjectAttach(SequencedDocumentMessageDescriptor descriptor, fluidDataStoreMessageAttach op);
    }

    /// <summary>
    /// IFluidDataObjectSender manages batch of changes to be broadcasted to other clients.
    /// caller has to pass an instance of IFluidDataObjectSender around when making model changes.
    /// all changes will be treated as "batch" and applied together.
    /// </summary>
    public interface IFluidDataObjectSender
    {
        /// <summary>
        /// Local client ID for the currently attached connection, or null before attach.
        /// </summary>
        string? LocalClientId => null;

        /// <summary>
        /// Queues an outbound op for a specific DDS. Returns the SequenceNumber
        /// stamped on the outgoing message (in particular the clientSequenceNumber
        /// used for pending-change tracking).
        /// </summary>
        /// <remarks>
        /// DEMO SIGNATURE: takes JSON directly. Host repo uses a Bond payload wrapping;
        /// reconcile on transfer (waccobalt owns the wire format).
        /// </remarks>
        SequenceNumber QueueDataObjectMessage(string address, string opTypeName, string opJson);
    }

    /// <summary>
    /// Resolves handles between DDS instances (URL ↔ object lookup).
    /// </summary>
    /// <remarks>
    /// Naming note: in Fluid TS the equivalent role is played by
    /// <c>IFluidHandleContext</c>. TS uses the word "registry" for something
    /// different — factory/type registration (<c>IFluidDataStoreRegistry</c>,
    /// <c>IChannelFactoryRegistry</c>) — so if you come from Fluid TS, ignore
    /// the name and read the members: this is the URL resolver, not a factory
    /// registry. The interface name matches waccobalt's real
    /// <c>IFluidDataObjectRegistry</c> exactly so the port drops straight in
    /// on transfer.
    /// </remarks>
    public interface IFluidDataObjectRegistry
    {
        IFluidDataObject? FindDataObject(string address);
        string GetDataObjectUrl(IFluidDataObject obj);
    }
}
