using System;
using System.DirectoryServices.AccountManagement;

namespace Microsoft.Dynamics.Nav.SMTP
{
	public static class MailHelpers
	{
		public static string TryGetEmailAddressFromActiveDirectory()
		{
			string emailAddress;
			try
			{
				emailAddress = UserPrincipal.Current.EmailAddress;
			}
			catch (InvalidOperationException invalidOperationException)
			{
				emailAddress = string.Empty;
			}
			catch (NoMatchingPrincipalException noMatchingPrincipalException)
			{
				emailAddress = string.Empty;
			}
			return emailAddress;
		}
	}
}