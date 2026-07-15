using System.Collections.Generic;

namespace Microsoft.Office.Web.Fluid.Tests
{
	internal sealed class FakeFluidDataObjectSender : IFluidDataObjectSender
	{
		private long _nextClientSeq = 1;

		public List<(string Address, string OpTypeName, string OpJson, long ClientSeq)> Sent { get; } = new();

		public SequenceNumber QueueDataObjectMessage(string address, string opTypeName, string opJson)
		{
			long clientSeq = _nextClientSeq++;
			Sent.Add((address, opTypeName, opJson, clientSeq));
			return SequenceNumber.ForTesting(clientSeq);
		}
	}
}
