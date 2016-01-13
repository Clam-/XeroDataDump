// https://github.com/XeroAPI/Xero-Net/blob/master/Xero.Api.Example.Applications/Private/Settings.cs

using System.Configuration;

namespace XeroDataDump
{
	public class Settings
	{
		public string Uri
		{
			get { return ConfigurationManager.AppSettings["BaseUrl"]; }
		}

		public string SigningCertificatePath
		{
			get { return ConfigurationManager.AppSettings["SigningCertificate"]; }
		}

		public string SigningCertificatePassword
		{
			get { return ConfigurationManager.AppSettings["SigningCertificatePassword"]; }
		}

		public string Key
		{
			get { return ConfigurationManager.AppSettings["ConsumerKey"]; }
		}

		public string Secret
		{
			get { return ConfigurationManager.AppSettings["ConsumerSecret"]; }
		}
	}
}
