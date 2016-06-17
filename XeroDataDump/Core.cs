using Xero.Api.Core;
using Xero.Api.Infrastructure.OAuth;
using Xero.Api.Infrastructure.RateLimiter;
using Xero.Api.Serialization;

namespace XeroDataDump
{
	public class Core : XeroCoreApi
	{
		private static readonly DefaultMapper Mapper = new DefaultMapper();

		public Core(bool includeRateLimiter = false) :
			base(Options.Default.BaseUrl,
				new PrivateAuthenticator(Options.Default.CertFile, Options.Default.CertPassword),
				new Consumer(Options.Default.ConsumerKey, Options.Default.ConsumerSecret),
				null,
				Mapper,
				Mapper,
				includeRateLimiter ? new RateLimiter() : null)
		{
		}
	}
}
