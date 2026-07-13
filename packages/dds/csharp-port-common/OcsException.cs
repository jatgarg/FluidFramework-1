// -----------------------------------------------------------------------------
// SHIM — deleted on transfer to host repo.
//
// Host repo defines OcsException and OcsGateErrorCode with a much larger set of
// error codes. We stub only what the DDS ports throw.
// -----------------------------------------------------------------------------

using System;

namespace Microsoft.Office.Web.Fluid
{
	public enum OcsGateErrorCode
	{
		UnknownOp,
		InvalidSequenceNumber,
		OutOfOrderSequenceNumber,
		InvalidOperation,
		InvalidState,
	}

	public class OcsException : Exception
	{
		public OcsGateErrorCode ErrorCode { get; }

		public OcsException(OcsGateErrorCode errorCode, string message)
			: base(message)
		{
			ErrorCode = errorCode;
		}

		public OcsException(OcsGateErrorCode errorCode, string message, Exception innerException)
			: base(message, innerException)
		{
			ErrorCode = errorCode;
		}
	}
}
