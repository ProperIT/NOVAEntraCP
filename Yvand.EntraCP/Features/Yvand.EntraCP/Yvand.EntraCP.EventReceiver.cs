using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using Microsoft.SharePoint.Administration.Claims;
using System;
using System.Runtime.InteropServices;
using Yvand.EntraClaimsProvider.Logging;

namespace Yvand.EntraClaimsProvider.Administration
{
    [Guid("39c10d12-2c7f-4148-bd81-2283a5ce4a27")]
    public class FarmFeatureEventReceiver : SPClaimProviderFeatureReceiver
    {
        private string _currentProviderName = EntraCP.DefaultClaimsProviderName;

        public override string ClaimProviderAssembly => typeof(EntraCP).Assembly.FullName;

        public override string ClaimProviderDescription => _currentProviderName;

        public override string ClaimProviderDisplayName => _currentProviderName;

        public override string ClaimProviderType => typeof(EntraCP).FullName;

        private static string[] GetProviderNames(SPFeatureReceiverProperties properties)
        {
            string names = properties.Feature?.Properties["ClaimProviderNames"]?.Value;
            if (String.IsNullOrWhiteSpace(names))
            {
                names = EntraCP.DefaultClaimsProviderName;
            }
            return names.Split(new[] { ';', ',' }, StringSplitOptions.RemoveEmptyEntries);
        }

        public override void FeatureActivated(SPFeatureReceiverProperties properties)
        {
            foreach (string name in GetProviderNames(properties))
            {
                _currentProviderName = name.Trim();
                SPSecurity.RunWithElevatedPrivileges(delegate ()
                {
                    try
                    {
                        Logger.Log($"[{_currentProviderName}] Activating farm-scoped feature for claims provider \"{_currentProviderName}\"", TraceSeverity.High, TraceCategory.Configuration);
                        base.FeatureActivated(properties);
                    }
                    catch (Exception ex)
                    {
                        Logger.LogException(_currentProviderName, $"activating farm-scoped feature for claims provider \"{_currentProviderName}\"", TraceCategory.Configuration, ex);
                    }
                });
            }
        }

        public override void FeatureUninstalling(SPFeatureReceiverProperties properties)
        {
            foreach (string name in GetProviderNames(properties))
            {
                SPSecurity.RunWithElevatedPrivileges(delegate ()
                {
                    try
                    {
                        Logger.Log($"[{name.Trim()}] Uninstalling farm-scoped feature for claims provider \"{name.Trim()}\": Deleting configuration from the farm", TraceSeverity.High, TraceCategory.Configuration);
                        Logger.Unregister();
                    }
                    catch (Exception ex)
                    {
                        Logger.LogException(name.Trim(), $"uninstalling farm-scoped feature for claims provider \"{name.Trim()}\"", TraceCategory.Configuration, ex);
                    }
                });
            }
        }

        public override void FeatureDeactivating(SPFeatureReceiverProperties properties)
        {
            foreach (string name in GetProviderNames(properties))
            {
                SPSecurity.RunWithElevatedPrivileges(delegate ()
                {
                    try
                    {
                        Logger.Log($"[{name.Trim()}] Deactivating farm-scoped feature for claims provider \"{name.Trim()}\": Removing claims provider from the farm (but not its configuration)", TraceSeverity.High, TraceCategory.Configuration);
                        base.RemoveClaimProvider(name.Trim());
                    }
                    catch (Exception ex)
                    {
                        Logger.LogException(name.Trim(), $"deactivating farm-scoped feature for claims provider \"{name.Trim()}\"", TraceCategory.Configuration, ex);
                    }
                });
            }
        }
    }
}
