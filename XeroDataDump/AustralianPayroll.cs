using Xero.Api.Infrastructure.OAuth;
using Xero.Api.Infrastructure.RateLimiter;
using Xero.Api.Serialization;

namespace XeroDataDump
{
	public class AustralianPayroll : Xero.Api.Payroll.AustralianPayroll
	{
		private static readonly DefaultMapper Mapper = new DefaultMapper();

		public AustralianPayroll(bool includeRateLimiter = false) :
			base(Options.Default.CallbackUrl,
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
