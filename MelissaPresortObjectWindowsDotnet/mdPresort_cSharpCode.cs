using System;
using System.Text;
using System.Runtime.InteropServices;
using System.Security;

namespace MelissaData {
	public class mdPresort : IDisposable {
		private IntPtr i;

		public enum ProgramStatus {
			ErrorNone = 0,
			ErrorOther = 1,
			ErrorOutOfMemory = 2,
			ErrorRequiredFileNotFound = 3,
			ErrorFoundOldFile = 4,
			ErrorDatabaseExpired = 5,
			ErrorLicenseExpired = 6
		}
		public enum AccessType {
			Local = 0,
			Remote = 1
		}
		public enum DiacriticsMode {
			Auto = 0,
			On = 1,
			Off = 2
		}
		public enum StandardizeMode {
			ShortFormat = 0,
			LongFormat = 1,
			AutoFormat = 2
		}
		public enum SuiteParseMode {
			ParseSuite = 0,
			CombineSuite = 1
		}
		public enum AliasPreserveMode {
			ConvertAlias = 0,
			PreserveAlias = 1
		}
		public enum AutoCompletionMode {
			AutoCompleteSingleSuite = 0,
			AutoCompleteRangedSuite = 1,
			AutoCompletePlaceHolderSuite = 2,
			AutoCompleteNoSuite = 3
		}
		public enum ResultCdDescOpt {
			ResultCodeDescriptionLong = 0,
			ResultCodeDescriptionShort = 1
		}
		public enum MailboxLookupMode {
			MailboxNone = 0,
			MailboxExpress = 1,
			MailboxPremium = 2
		}
		public enum SortationCode {
			FCM_LTR_AUTO_NONAUTO = 1,
			FCM_LTR_AUTO = 2,
			FCM_LTR_NONAUTO = 3,
			FCM_LTR_NONMACH = 4,
			FCM_POSTCARD_AUTO_NONAUTO = 41,
			FCM_POSTCARD_AUTO = 42,
			FCM_POSTCARD_NONAUTO = 43,
			FCM_FLAT_COTRAY = 51,
			FCM_FLAT_AUTO = 52,
			FCM_FLAT_NONAUTO = 53,
			FCM_FLAT_DISABLE_COTRAY = 54,
			FCM_FLAT_COTRAY_FSM1000 = 55,
			FCM_FLAT_DISABLE_COTRAY_FSM1000 = 56,
			FCM_FLAT_AUTO_FSM1000 = 57,
			FCM_FLAT_NONAUTO_FSM1000 = 58,
			STD_LTR_ECRRT_AUTO_NONAUTO = 101,
			STD_LTR_ECRRT_NONAUTO = 102,
			STD_LTR_AUTO_NONAUTO = 103,
			STD_LTR_AUTO = 104,
			STD_LTR_NONAUTO = 105,
			STD_LTR_NONMACH = 106,
			STD_LTR_ECRRT_NONMACH = 107,
			STD_LTR_ECRRT_AUTO = 108,
			STD_LTR_ECRRT = 109,
			STD_FLAT_ECRRT_COSACK = 151,
			STD_FLAT_ECRRT_DISABLE_COSACK = 152,
			STD_FLAT_COSACK = 153,
			STD_FLAT_DISABLE_COSACK = 154,
			STD_FLAT_AUTO = 155,
			STD_FLAT_NONAUTO = 156,
			STD_FLAT_ECRRT_AUTO = 157,
			STD_FLAT_ECRRT_NONAUTO = 158,
			STD_FLAT_ECRRT = 159,
			RESIDUALS_LEFT = 999
		}
		public enum ACSCode {
			FCM_ACS_ASR = 1,
			FCM_ACS_ASR2 = 2,
			FCM_ACS_CSR = 3,
			FCM_ACS_CSR2 = 4,
			FCM_ACS_ASR_ONECODE = 5,
			FCM_ACS_ASR2_ONECODE = 6,
			FCM_ACS_CSR_ONECODE = 7,
			FCM_ACS_CSR2_ONECODE = 8,
			FCM_MANUAL_CORRECTION = 9,
			FCM_NO_ADDR_CORRECTION = 10,
			FCM_ACS_ASR_TRACE = 11,
			FCM_ACS_ASR2_TRACE = 12,
			FCM_ACS_CSR_TRACE = 13,
			FCM_ACS_CSR2_TRACE = 14,
			FCM_ACS_ASR_ONECODE_TRACE = 15,
			FCM_ACS_CSR_ONECODE_TRACE = 16,
			FCM_MANUAL_CORRECTION_TRACE = 17,
			FCM_NO_ADDR_CORRECTION_TRACE = 18,
			FCM_ACS_ASR_FS = 19,
			FCM_ACS_ASR2_FS = 20,
			FCM_ACS_CSR_FS = 21,
			FCM_ACS_CSR2_FS = 22,
			FCM_FULL_SERVICE_ACS_ASR_FS = 23,
			FCM_FULL_SERVICE_ACS_ASR2_FS = 24,
			FCM_MANUAL_CORRECTION_FS = 25,
			FCM_NO_ADDR_CORRECTION_FS = 26,
			FCM_ACS_ASR_FS_TRACE = 27,
			FCM_ACS_ASR2_FS_TRACE = 28,
			FCM_ACS_CSR_FS_TRACE = 29,
			FCM_ACS_CSR2_FS_TRACE = 30,
			FCM_FULL_SERVICE_ACS_ASR_FS_TRACE = 31,
			FCM_FULL_SERVICE_ACS_ASR2_FS_TRACE = 32,
			FCM_MANUAL_CORRECTION_FS_TRACE = 33,
			FCM_NO_ADDR_CORRECTION_FS_TRACE = 34,
			FCM_ACS_ASR2_ONECODE_TRACE = 35,
			FCM_ACS_CSR2_ONECODE_TRACE = 36,
			FCM_FULL_SERVICE_ACS_CSR_FS = 37,
			FCM_FULL_SERVICE_ACS_CSR2_FS = 38,
			FCM_FULL_SERVICE_ACS_CSR_FS_TRACE = 39,
			FCM_FULL_SERVICE_ACS_CSR2_FS_TRACE = 40,
			FCM_ACS_RSR2 = 41,
			FCM_ACS_TRSR2 = 42,
			FCM_ACS_RSR2_ONECODE = 43,
			FCM_ACS_TRSR2_ONECODE = 44,
			FCM_ACS_RSR2_TRACE = 45,
			FCM_ACS_TRSR2_TRACE = 46,
			FCM_ACS_RSR2_ONECODE_TRACE = 47,
			FCM_ACS_TRSR2_ONECODE_TRACE = 48,
			FCM_ACS_RSR2_FS = 49,
			FCM_ACS_TRSR2_FS = 50,
			FCM_FULL_SERVICE_ACS_RSR2_FS = 51,
			FCM_FULL_SERVICE_ACS_TRSR2_FS = 52,
			FCM_ACS_RSR2_FS_TRACE = 53,
			FCM_ACS_TRSR2_FS_TRACE = 54,
			FCM_FULL_SERVICE_ACS_RSR2_FS_TRACE = 55,
			FCM_FULL_SERVICE_ACS_TRSR2_FS_TRACEFCM_ACS_CSR_ONECODE_SD = 56,
			FCM_ACS_CSR_ONECODE_SD_IV_MTR = 57,
			FCM_ACS_CSR2_ONECODE_SD = 58,
			FCM_ACS_CSR2_ONECODE_SD_IV_MTR = 59,
			FCM_FULL_SERVICE_CSR_SD_FS = 60,
			FCM_FULL_SERVICE_CSR_SD_IV_MTR_FS = 61,
			FCM_FULL_SERVICE_CSR2_SD_FS = 62,
			FCM_FULL_SERVICE_CSR2_SD_IV_MTR_FS = 63,
			FCM_ACS_CSR_SD = 64,
			FCM_ACS_CSR_SD_IV_MTR = 65,
			FCM_ACS_CSR2_SD = 66,
			FCM_ACS_CSR2_IV_MTR_SD = 67,
			FCM_ACS_CSR_SD_FS = 68,
			FCM_ACS_CSR_SD_IV_MTR_FS = 69,
			FCM_ACS_CSR2_SD_FS = 70,
			FCM_ACS_CSR2_IV_MTR_SD_FS = 71,
			STD_ACS_ASR = 101,
			STD_ACS_CSR = 102,
			STD_ACS_ASR_ONECODE = 103,
			STD_ACS_CSR_ONECODE = 104,
			STD_MANUAL_CORRECTION = 105,
			STD_NO_ADDR_CORRECTION = 106,
			STD_ACS_ASR_TRACE = 107,
			STD_ACS_CSR_TRACE = 108,
			STD_ACS_ASR_ONECODE_TRACE = 109,
			STD_ACS_CSR_ONECODE_TRACE = 110,
			STD_MANUAL_CORRECTION_TRACE = 111,
			STD_NO_ADDR_CORRECTION_TRACE = 112,
			STD_ACS_ASR_FS = 113,
			STD_ACS_CSR_FS = 114,
			STD_FULL_SERVICE_ACS_ASR_FS = 115,
			STD_FULL_SERVICE_ACS_CSR_FS = 116,
			STD_MANUAL_CORRECTION_FS = 117,
			STD_NO_ADDR_CORRECTION_FS = 118,
			STD_ACS_ASR_FS_TRACE = 119,
			STD_ACS_CSR_FS_TRACE = 120,
			STD_FULL_SERVICE_ACS_ASR_FS_TRACE = 121,
			STD_FULL_SERVICE_ACS_CSR_FS_TRACE = 122,
			STD_MANUAL_CORRECTION_FS_TRACE = 123,
			STD_NO_ADDR_CORRECTION_FS_TRACE = 124,
			STD_ACS_ASR2 = 125,
			STD_ACS_RSR2 = 126,
			STD_ACS_ASR2_ONECODE = 127,
			STD_ACS_RSR2_ONECODE = 128,
			STD_ACS_ASR2_TRACE = 129,
			STD_ACS_RSR2_TRACE = 130,
			STD_ACS_ASR2_ONECODE_TRACE = 131,
			STD_ACS_RSR2_ONECODE_TRACE = 132,
			STD_ACS_ASR2_FS = 133,
			STD_ACS_RSR2_FS = 134,
			STD_FULL_SERVICE_ACS_ASR2_FS = 135,
			STD_FULL_SERVICE_ACS_RSR2_FS = 136,
			STD_ACS_ASR2_FS_TRACE = 137,
			STD_ACS_RSR2_FS_TRACE = 138,
			STD_FULL_SERVICE_ACS_ASR2_FS_TRACE = 139,
			STD_FULL_SERVICE_ACS_RSR2_FS_TRACE = 140,
			STD_ACS_CSR2 = 141,
			STD_ACS_CSR2_TRACE = 142,
			STD_ACS_CSR2_ONECODE = 143,
			STD_ACS_CSR2_ONECODE_TRACE = 144,
			STD_ACS_CSR2_FS = 145,
			STD_ACS_CSR2_FS_TRACE = 146,
			STD_FULL_SERVICE_ACS_CSR2_FS = 147,
			STD_FULL_SERVICE_ACS_CSR2_FS_TRACE = 148,
			ACS_LAST = 999
		}

		[SuppressUnmanagedCodeSecurity]
		private class mdPresortUnmanaged {
			[DllImport("mdPresort", EntryPoint = "mdPresortCreate", CallingConvention = CallingConvention.Cdecl)]
			public static extern IntPtr mdPresortCreate();
			[DllImport("mdPresort", EntryPoint = "mdPresortDestroy", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortDestroy(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortInitializeDataFiles", CallingConvention = CallingConvention.Cdecl)]
			public static extern Int32 mdPresortInitializeDataFiles(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetLicenseString", CallingConvention = CallingConvention.Cdecl)]
			public static extern Int32 mdPresortSetLicenseString(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetPathToPresortDataFiles", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetPathToPresortDataFiles(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetInitializeErrorString", CallingConvention = CallingConvention.Cdecl)]
			public static extern IntPtr mdPresortGetInitializeErrorString(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetParametersErrorString", CallingConvention = CallingConvention.Cdecl)]
			public static extern IntPtr mdPresortGetParametersErrorString(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetBuildNumber", CallingConvention = CallingConvention.Cdecl)]
			public static extern IntPtr mdPresortGetBuildNumber(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetLicenseStringExpirationDate", CallingConvention = CallingConvention.Cdecl)]
			public static extern IntPtr mdPresortGetLicenseStringExpirationDate(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetDatabaseExpirationDate", CallingConvention = CallingConvention.Cdecl)]
			public static extern IntPtr mdPresortGetDatabaseExpirationDate(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetDatabaseDate", CallingConvention = CallingConvention.Cdecl)]
			public static extern IntPtr mdPresortGetDatabaseDate(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetZip", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetZip(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetPlus4", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetPlus4(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetCarrierRoute", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetCarrierRoute(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetWalkSequence", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetWalkSequence(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetDeliveryPointCode", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetDeliveryPointCode(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetBusinessResidentialIndicator", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetBusinessResidentialIndicator(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetPSPermitImprint", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetPSPermitImprint(IntPtr i, Int32 p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetLOTNumber", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetLOTNumber(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetLOTOrder", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetLOTOrder(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetPSASE", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetPSASE(IntPtr i, Int32 p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetPSFASTforward", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetPSFASTforward(IntPtr i, Int32 p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetPSNCOA", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetPSNCOA(IntPtr i, Int32 p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetPSACS", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetPSACS(IntPtr i, Int32 p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetPSAltMethod", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetPSAltMethod(IntPtr i, Int32 p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetPSMultiple", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetPSMultiple(IntPtr i, Int32 p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetPSOneCode", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetPSOneCode(IntPtr i, Int32 p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetPSAltAddFmt", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetPSAltAddFmt(IntPtr i, Int32 p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetPSCASSDate", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetPSCASSDate(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetPSPrecanceledStamp", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetPSPrecanceledStamp(IntPtr i, Int32 p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetPSPermitHolderName", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetPSPermitHolderName(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetPSPermitHolderCompany", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetPSPermitHolderCompany(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetPSPermitHolderStreet", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetPSPermitHolderStreet(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetPSPermitHolderCity", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetPSPermitHolderCity(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetPSPermitHolderState", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetPSPermitHolderState(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetPSPermitHolderPhone", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetPSPermitHolderPhone(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetPSPermitHolderEmail", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetPSPermitHolderEmail(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetPSPermitHolderListName", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetPSPermitHolderListName(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetPSPermitHolderZIP", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetPSPermitHolderZIP(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetPSPostOfficeOfMailingCity", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetPSPostOfficeOfMailingCity(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetPSPostOfficeOfMailingState", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetPSPostOfficeOfMailingState(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetPSPostOfficeOfMailingZIP", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetPSPostOfficeOfMailingZIP(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetPSMailingAgentName", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetPSMailingAgentName(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetPSMailingAgentCompany", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetPSMailingAgentCompany(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetPSMailingAgentStreet", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetPSMailingAgentStreet(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetPSMailingAgentCity", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetPSMailingAgentCity(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetPSMailingAgentState", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetPSMailingAgentState(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetPSMailingAgentPhone", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetPSMailingAgentPhone(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetPSMailingAgentZIP", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetPSMailingAgentZIP(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetPSIndividualOrOrganizationName", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetPSIndividualOrOrganizationName(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetPSIndividualOrOrganizationCompany", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetPSIndividualOrOrganizationCompany(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetPSIndividualOrOrganizationStreet", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetPSIndividualOrOrganizationStreet(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetPSIndividualOrOrganizationCity", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetPSIndividualOrOrganizationCity(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetPSIndividualOrOrganizationState", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetPSIndividualOrOrganizationState(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetPSIndividualOrOrganizationZIP", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetPSIndividualOrOrganizationZIP(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetPSIndividualOrOrganizationCRID", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetPSIndividualOrOrganizationCRID(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetContinueContainerNumber", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetContinueContainerNumber(IntPtr i, Int32 p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetPOMasNDC", CallingConvention = CallingConvention.Cdecl)]
			public static extern Int32 mdPresortSetPOMasNDC(IntPtr i, Int32 p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetPOMasSCF", CallingConvention = CallingConvention.Cdecl)]
			public static extern Int32 mdPresortSetPOMasSCF(IntPtr i, Int32 p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetPSStatementSeqNumber", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetPSStatementSeqNumber(IntPtr i, IntPtr val);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetPSMailingDate", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetPSMailingDate(IntPtr i, IntPtr val);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetPSFedAgencyCode", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetPSFedAgencyCode(IntPtr i, IntPtr val);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetPSCustomerNumber", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetPSCustomerNumber(IntPtr i, IntPtr val);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetPSCAPSNumber", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetPSCAPSNumber(IntPtr i, IntPtr val);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetPermitNumber", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetPermitNumber(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetPSNonProfitAuthNumber", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetPSNonProfitAuthNumber(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetMailClass", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetMailClass(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetPieceType", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetPieceType(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetSortType", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetSortType(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetFCM_CRRT", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetFCM_CRRT(IntPtr i, Int32 p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetFCM_CRRT_5", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetFCM_CRRT_5(IntPtr i, Int32 p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetFCM_CRRT_3", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetFCM_CRRT_3(IntPtr i, Int32 p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetFCM_5dg_Scheme", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetFCM_5dg_Scheme(IntPtr i, Int32 p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetFCM_Auto_5dg", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetFCM_Auto_5dg(IntPtr i, Int32 p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetFCM_NonAuto_5dg", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetFCM_NonAuto_5dg(IntPtr i, Int32 p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetSTD_Auto_5dg_Scheme", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetSTD_Auto_5dg_Scheme(IntPtr i, Int32 p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetSTD_Auto_5dg", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetSTD_Auto_5dg(IntPtr i, Int32 p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetNonMachinableType", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetNonMachinableType(IntPtr i, Int32 p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetRecordID", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetRecordID(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetSackWeight", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetSackWeight(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetPieceLength", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetPieceLength(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetPieceHeight", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetPieceHeight(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetPieceThickness", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetPieceThickness(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetPieceWeight", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetPieceWeight(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetPackageSize", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetPackageSize(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetEnableCoTray", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetEnableCoTray(IntPtr i, Int32 p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetUseFSM1000", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetUseFSM1000(IntPtr i, Int32 p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetSTDNonProfit", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetSTDNonProfit(IntPtr i, Int32 p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetIgnoreDSF", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetIgnoreDSF(IntPtr i, Int32 val);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetMailersID", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetMailersID(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetPSPrecanceledStampValue", CallingConvention = CallingConvention.Cdecl)]
			public static extern Int32 mdPresortSetPSPrecanceledStampValue(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetPSPresortResidualPieces", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetPSPresortResidualPieces(IntPtr i, Int32 p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetIMBSerialNumber", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetIMBSerialNumber(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetPSPolicitalCampaignMailing", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetPSPolicitalCampaignMailing(IntPtr i, Int32 p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetPSOfficialElectionMail", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetPSOfficialElectionMail(IntPtr i, Int32 p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortProduceReports", CallingConvention = CallingConvention.Cdecl)]
			public static extern Int32 mdPresortProduceReports(IntPtr i, IntPtr p1, IntPtr p2);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetSCFCity", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetSCFCity(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetSCFState", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetSCFState(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetSCFZip", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetSCFZip(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortAddSCF", CallingConvention = CallingConvention.Cdecl)]
			public static extern Int32 mdPresortAddSCF(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetNDCCity", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetNDCCity(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetNDCState", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetNDCState(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetNDCZip", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetNDCZip(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortAddNDC", CallingConvention = CallingConvention.Cdecl)]
			public static extern Int32 mdPresortAddNDC(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetDDUCity", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetDDUCity(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetDDUState", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetDDUState(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetDDUZip", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetDDUZip(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetDDUMoreZip", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetDDUMoreZip(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortAddDDU", CallingConvention = CallingConvention.Cdecl)]
			public static extern Int32 mdPresortAddDDU(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetTTBorder", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetTTBorder(IntPtr i, Int32 p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetTTNumberofPieces", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetTTNumberofPieces(IntPtr i, Int32 p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetTTContainerNumber", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetTTContainerNumber(IntPtr i, Int32 val);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetTTContainerSize", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetTTContainerSize(IntPtr i, Int32 val);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetTTOther", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetTTOther(IntPtr i, IntPtr val2);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetTTParameterPositionX", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetTTParameterPositionX(IntPtr i, IntPtr val);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetTTParameterPositionY", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetTTParameterPositionY(IntPtr i, IntPtr val);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetTTParameterWidth", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetTTParameterWidth(IntPtr i, IntPtr val);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetTTParameterHeight", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetTTParameterHeight(IntPtr i, IntPtr val);
			[DllImport("mdPresort", EntryPoint = "mdPresortDoPresort", CallingConvention = CallingConvention.Cdecl)]
			public static extern Int32 mdPresortDoPresort(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetDisableDetailedOutput", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetDisableDetailedOutput(IntPtr i, Int32 p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetPreSortSettings", CallingConvention = CallingConvention.Cdecl)]
			public static extern IntPtr mdPresortSetPreSortSettings(IntPtr i, Int32 p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetACSCodeSettings", CallingConvention = CallingConvention.Cdecl)]
			public static extern Int32 mdPresortSetACSCodeSettings(IntPtr i, Int32 p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortUpdateParameters", CallingConvention = CallingConvention.Cdecl)]
			public static extern Int32 mdPresortUpdateParameters(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortAddRecord", CallingConvention = CallingConvention.Cdecl)]
			public static extern Int32 mdPresortAddRecord(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetRecord", CallingConvention = CallingConvention.Cdecl)]
			public static extern Int32 mdPresortGetRecord(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetFirstRecord", CallingConvention = CallingConvention.Cdecl)]
			public static extern Int32 mdPresortGetFirstRecord(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetNextRecord", CallingConvention = CallingConvention.Cdecl)]
			public static extern Int32 mdPresortGetNextRecord(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetMailPieceRate", CallingConvention = CallingConvention.Cdecl)]
			public static extern IntPtr mdPresortGetMailPieceRate(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetZipAsString", CallingConvention = CallingConvention.Cdecl)]
			public static extern IntPtr mdPresortGetZipAsString(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetTrayNumber", CallingConvention = CallingConvention.Cdecl)]
			public static extern Int32 mdPresortGetTrayNumber(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetSequenceNumber", CallingConvention = CallingConvention.Cdecl)]
			public static extern Int32 mdPresortGetSequenceNumber(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetRecordID", CallingConvention = CallingConvention.Cdecl)]
			public static extern IntPtr mdPresortGetRecordID(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetEndorsementLine", CallingConvention = CallingConvention.Cdecl)]
			public static extern IntPtr mdPresortGetEndorsementLine(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetRateCode", CallingConvention = CallingConvention.Cdecl)]
			public static extern IntPtr mdPresortGetRateCode(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetMailJob", CallingConvention = CallingConvention.Cdecl)]
			public static extern IntPtr mdPresortGetMailJob(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetSortLevel", CallingConvention = CallingConvention.Cdecl)]
			public static extern IntPtr mdPresortGetSortLevel(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetLabelLine1", CallingConvention = CallingConvention.Cdecl)]
			public static extern IntPtr mdPresortGetLabelLine1(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetLabelLine2", CallingConvention = CallingConvention.Cdecl)]
			public static extern IntPtr mdPresortGetLabelLine2(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetLabelList", CallingConvention = CallingConvention.Cdecl)]
			public static extern IntPtr mdPresortGetLabelList(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetCINCode", CallingConvention = CallingConvention.Cdecl)]
			public static extern IntPtr mdPresortGetCINCode(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetTrayProcessingCode", CallingConvention = CallingConvention.Cdecl)]
			public static extern IntPtr mdPresortGetTrayProcessingCode(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetMailSplitCode", CallingConvention = CallingConvention.Cdecl)]
			public static extern IntPtr mdPresortGetMailSplitCode(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetTrayZipCode", CallingConvention = CallingConvention.Cdecl)]
			public static extern IntPtr mdPresortGetTrayZipCode(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetTrayType", CallingConvention = CallingConvention.Cdecl)]
			public static extern IntPtr mdPresortGetTrayType(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetBundleZipCode", CallingConvention = CallingConvention.Cdecl)]
			public static extern IntPtr mdPresortGetBundleZipCode(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetBundleSortLevel", CallingConvention = CallingConvention.Cdecl)]
			public static extern IntPtr mdPresortGetBundleSortLevel(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetBundleNumber", CallingConvention = CallingConvention.Cdecl)]
			public static extern Int32 mdPresortGetBundleNumber(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetBundleIndicator", CallingConvention = CallingConvention.Cdecl)]
			public static extern IntPtr mdPresortGetBundleIndicator(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetZipCodeInScheme", CallingConvention = CallingConvention.Cdecl)]
			public static extern IntPtr mdPresortGetZipCodeInScheme(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetBarcodeID", CallingConvention = CallingConvention.Cdecl)]
			public static extern IntPtr mdPresortGetBarcodeID(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetServiceTypeID", CallingConvention = CallingConvention.Cdecl)]
			public static extern IntPtr mdPresortGetServiceTypeID(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetProducePallets", CallingConvention = CallingConvention.Cdecl)]
			public static extern Int32 mdPresortSetProducePallets(IntPtr i, Int32 p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetIgnoreMinSCFPallet", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetIgnoreMinSCFPallet(IntPtr i, Int32 val);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetPalletsTotal", CallingConvention = CallingConvention.Cdecl)]
			public static extern Int32 mdPresortGetPalletsTotal(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetPalletZipCode", CallingConvention = CallingConvention.Cdecl)]
			public static extern IntPtr mdPresortGetPalletZipCode(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetPalletLabelLine1", CallingConvention = CallingConvention.Cdecl)]
			public static extern IntPtr mdPresortGetPalletLabelLine1(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetPalletLabelLine2", CallingConvention = CallingConvention.Cdecl)]
			public static extern IntPtr mdPresortGetPalletLabelLine2(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetPalletSortLevel", CallingConvention = CallingConvention.Cdecl)]
			public static extern IntPtr mdPresortGetPalletSortLevel(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetPalletNumber", CallingConvention = CallingConvention.Cdecl)]
			public static extern Int32 mdPresortGetPalletNumber(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetTotalPiecesAuto5dig", CallingConvention = CallingConvention.Cdecl)]
			public static extern Int32 mdPresortGetTotalPiecesAuto5dig(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetTotalPiecesAuto3dig", CallingConvention = CallingConvention.Cdecl)]
			public static extern Int32 mdPresortGetTotalPiecesAuto3dig(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetTotalPiecesAutoADC", CallingConvention = CallingConvention.Cdecl)]
			public static extern Int32 mdPresortGetTotalPiecesAutoADC(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetTotalPiecesAutoMADC", CallingConvention = CallingConvention.Cdecl)]
			public static extern Int32 mdPresortGetTotalPiecesAutoMADC(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetTotalPieces5dig", CallingConvention = CallingConvention.Cdecl)]
			public static extern Int32 mdPresortGetTotalPieces5dig(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetTotalPieces3dig", CallingConvention = CallingConvention.Cdecl)]
			public static extern Int32 mdPresortGetTotalPieces3dig(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetTotalPiecesADC", CallingConvention = CallingConvention.Cdecl)]
			public static extern Int32 mdPresortGetTotalPiecesADC(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetTotalPiecesMADC", CallingConvention = CallingConvention.Cdecl)]
			public static extern Int32 mdPresortGetTotalPiecesMADC(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetTotalPiecesWS", CallingConvention = CallingConvention.Cdecl)]
			public static extern Int32 mdPresortGetTotalPiecesWS(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetTotalPiecesHDP", CallingConvention = CallingConvention.Cdecl)]
			public static extern Int32 mdPresortGetTotalPiecesHDP(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetTotalPiecesHD", CallingConvention = CallingConvention.Cdecl)]
			public static extern Int32 mdPresortGetTotalPiecesHD(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetTotalPiecesLOT", CallingConvention = CallingConvention.Cdecl)]
			public static extern Int32 mdPresortGetTotalPiecesLOT(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetRateAuto5dig", CallingConvention = CallingConvention.Cdecl)]
			public static extern double mdPresortGetRateAuto5dig(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetRateAuto3dig", CallingConvention = CallingConvention.Cdecl)]
			public static extern double mdPresortGetRateAuto3dig(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetRateAutoADC", CallingConvention = CallingConvention.Cdecl)]
			public static extern double mdPresortGetRateAutoADC(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetRateAutoMADC", CallingConvention = CallingConvention.Cdecl)]
			public static extern double mdPresortGetRateAutoMADC(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetRate5dig", CallingConvention = CallingConvention.Cdecl)]
			public static extern double mdPresortGetRate5dig(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetRate3dig", CallingConvention = CallingConvention.Cdecl)]
			public static extern double mdPresortGetRate3dig(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetRateADC", CallingConvention = CallingConvention.Cdecl)]
			public static extern double mdPresortGetRateADC(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetRateMADC", CallingConvention = CallingConvention.Cdecl)]
			public static extern double mdPresortGetRateMADC(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetRateWS", CallingConvention = CallingConvention.Cdecl)]
			public static extern double mdPresortGetRateWS(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetRateHDP", CallingConvention = CallingConvention.Cdecl)]
			public static extern double mdPresortGetRateHDP(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetRateHD", CallingConvention = CallingConvention.Cdecl)]
			public static extern double mdPresortGetRateHD(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetRateLOT", CallingConvention = CallingConvention.Cdecl)]
			public static extern double mdPresortGetRateLOT(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetPoundRate", CallingConvention = CallingConvention.Cdecl)]
			public static extern double mdPresortGetPoundRate(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetPoundRateWS", CallingConvention = CallingConvention.Cdecl)]
			public static extern double mdPresortGetPoundRateWS(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetPoundRateHDP", CallingConvention = CallingConvention.Cdecl)]
			public static extern double mdPresortGetPoundRateHDP(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetPoundRateHD", CallingConvention = CallingConvention.Cdecl)]
			public static extern double mdPresortGetPoundRateHD(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetPoundRateLOT", CallingConvention = CallingConvention.Cdecl)]
			public static extern double mdPresortGetPoundRateLOT(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetTotalTrays1Ft", CallingConvention = CallingConvention.Cdecl)]
			public static extern Int32 mdPresortGetTotalTrays1Ft(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetTotalTrays2Ft", CallingConvention = CallingConvention.Cdecl)]
			public static extern Int32 mdPresortGetTotalTrays2Ft(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetTotalSacks", CallingConvention = CallingConvention.Cdecl)]
			public static extern Int32 mdPresortGetTotalSacks(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetTotalResidualPieces", CallingConvention = CallingConvention.Cdecl)]
			public static extern Int32 mdPresortGetTotalResidualPieces(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetTotalPiecesAuto5digFSS", CallingConvention = CallingConvention.Cdecl)]
			public static extern Int32 mdPresortGetTotalPiecesAuto5digFSS(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetTotalPieces5digFSS", CallingConvention = CallingConvention.Cdecl)]
			public static extern Int32 mdPresortGetTotalPieces5digFSS(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetRateAuto5digFSS", CallingConvention = CallingConvention.Cdecl)]
			public static extern double mdPresortGetRateAuto5digFSS(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetRate5digFSS", CallingConvention = CallingConvention.Cdecl)]
			public static extern double mdPresortGetRate5digFSS(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetPoundRateAuto5digFSS", CallingConvention = CallingConvention.Cdecl)]
			public static extern double mdPresortGetPoundRateAuto5digFSS(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetPoundRate5digFSS", CallingConvention = CallingConvention.Cdecl)]
			public static extern double mdPresortGetPoundRate5digFSS(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetDropShipZipPlus4", CallingConvention = CallingConvention.Cdecl)]
			public static extern IntPtr mdPresortGetDropShipZipPlus4(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGet", CallingConvention = CallingConvention.Cdecl)]
			public static extern IntPtr mdPresortGet(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetTotalNonAutoPieces", CallingConvention = CallingConvention.Cdecl)]
			public static extern Int32 mdPresortGetTotalNonAutoPieces(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetPresortedRate", CallingConvention = CallingConvention.Cdecl)]
			public static extern double mdPresortGetPresortedRate(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetResidualRate", CallingConvention = CallingConvention.Cdecl)]
			public static extern double mdPresortGetResidualRate(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetDestTotalPiecesAuto5dig", CallingConvention = CallingConvention.Cdecl)]
			public static extern Int32 mdPresortGetDestTotalPiecesAuto5dig(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetDestTotalPiecesAuto3dig", CallingConvention = CallingConvention.Cdecl)]
			public static extern Int32 mdPresortGetDestTotalPiecesAuto3dig(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetDestTotalPiecesAutoADC", CallingConvention = CallingConvention.Cdecl)]
			public static extern Int32 mdPresortGetDestTotalPiecesAutoADC(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetDestTotalPiecesAutoMADC", CallingConvention = CallingConvention.Cdecl)]
			public static extern Int32 mdPresortGetDestTotalPiecesAutoMADC(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetDestTotalPieces5dig", CallingConvention = CallingConvention.Cdecl)]
			public static extern Int32 mdPresortGetDestTotalPieces5dig(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetDestTotalPieces3dig", CallingConvention = CallingConvention.Cdecl)]
			public static extern Int32 mdPresortGetDestTotalPieces3dig(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetDestTotalPiecesADC", CallingConvention = CallingConvention.Cdecl)]
			public static extern Int32 mdPresortGetDestTotalPiecesADC(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetDestTotalPiecesMADC", CallingConvention = CallingConvention.Cdecl)]
			public static extern Int32 mdPresortGetDestTotalPiecesMADC(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetDestTotalPiecesWS", CallingConvention = CallingConvention.Cdecl)]
			public static extern Int32 mdPresortGetDestTotalPiecesWS(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetDestTotalPiecesHDP", CallingConvention = CallingConvention.Cdecl)]
			public static extern Int32 mdPresortGetDestTotalPiecesHDP(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetDestTotalPiecesHD", CallingConvention = CallingConvention.Cdecl)]
			public static extern Int32 mdPresortGetDestTotalPiecesHD(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetDestTotalPiecesLOT", CallingConvention = CallingConvention.Cdecl)]
			public static extern Int32 mdPresortGetDestTotalPiecesLOT(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetDestRateAuto5dig", CallingConvention = CallingConvention.Cdecl)]
			public static extern double mdPresortGetDestRateAuto5dig(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetDestRateAuto3dig", CallingConvention = CallingConvention.Cdecl)]
			public static extern double mdPresortGetDestRateAuto3dig(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetDestRateAutoADC", CallingConvention = CallingConvention.Cdecl)]
			public static extern double mdPresortGetDestRateAutoADC(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetDestRateAutoMADC", CallingConvention = CallingConvention.Cdecl)]
			public static extern double mdPresortGetDestRateAutoMADC(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetDestRate5dig", CallingConvention = CallingConvention.Cdecl)]
			public static extern double mdPresortGetDestRate5dig(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetDestRate3dig", CallingConvention = CallingConvention.Cdecl)]
			public static extern double mdPresortGetDestRate3dig(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetDestRateADC", CallingConvention = CallingConvention.Cdecl)]
			public static extern double mdPresortGetDestRateADC(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetDestRateMADC", CallingConvention = CallingConvention.Cdecl)]
			public static extern double mdPresortGetDestRateMADC(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetDestRateWS", CallingConvention = CallingConvention.Cdecl)]
			public static extern double mdPresortGetDestRateWS(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetDestRateHDP", CallingConvention = CallingConvention.Cdecl)]
			public static extern double mdPresortGetDestRateHDP(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetDestRateHD", CallingConvention = CallingConvention.Cdecl)]
			public static extern double mdPresortGetDestRateHD(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetDestRateLOT", CallingConvention = CallingConvention.Cdecl)]
			public static extern double mdPresortGetDestRateLOT(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetContainersTotal", CallingConvention = CallingConvention.Cdecl)]
			public static extern Int32 mdPresortGetContainersTotal(IntPtr i, Int32 p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetBundlesTotal", CallingConvention = CallingConvention.Cdecl)]
			public static extern Int32 mdPresortGetBundlesTotal(IntPtr i, Int32 p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetPiecesTotal", CallingConvention = CallingConvention.Cdecl)]
			public static extern Int32 mdPresortGetPiecesTotal(IntPtr i, Int32 p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetChargeTotal", CallingConvention = CallingConvention.Cdecl)]
			public static extern double mdPresortGetChargeTotal(IntPtr i, Int32 p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetDestSCFInfo", CallingConvention = CallingConvention.Cdecl)]
			public static extern Int32 mdPresortGetDestSCFInfo(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetDestNDCInfo", CallingConvention = CallingConvention.Cdecl)]
			public static extern Int32 mdPresortGetDestNDCInfo(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetDestDDUInfo", CallingConvention = CallingConvention.Cdecl)]
			public static extern Int32 mdPresortGetDestDDUInfo(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetDestContainersTotal", CallingConvention = CallingConvention.Cdecl)]
			public static extern Int32 mdPresortGetDestContainersTotal(IntPtr i, Int32 p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetDestBundlesTotal", CallingConvention = CallingConvention.Cdecl)]
			public static extern Int32 mdPresortGetDestBundlesTotal(IntPtr i, Int32 p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetDestPiecesTotal", CallingConvention = CallingConvention.Cdecl)]
			public static extern Int32 mdPresortGetDestPiecesTotal(IntPtr i, Int32 p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetDestChargeTotal", CallingConvention = CallingConvention.Cdecl)]
			public static extern double mdPresortGetDestChargeTotal(IntPtr i, Int32 p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetTwoFtTrayMaximum", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetTwoFtTrayMaximum(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetTwoFtTrayMinimum", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetTwoFtTrayMinimum(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetOneFtTrayMaximum", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetOneFtTrayMaximum(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetOneFtTrayMinimum", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetOneFtTrayMinimum(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetBundleMaximum", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetBundleMaximum(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetPSSCFZipAutomationSort", CallingConvention = CallingConvention.Cdecl)]
			public static extern Int32 mdPresortSetPSSCFZipAutomationSort(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetPSSCFZipNonAutomationSort", CallingConvention = CallingConvention.Cdecl)]
			public static extern Int32 mdPresortSetPSSCFZipNonAutomationSort(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetPSSCFZipECRRTSort", CallingConvention = CallingConvention.Cdecl)]
			public static extern Int32 mdPresortSetPSSCFZipECRRTSort(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetPSSCFZipCoSack", CallingConvention = CallingConvention.Cdecl)]
			public static extern Int32 mdPresortSetPSSCFZipCoSack(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetPSNDCZipAutomationSort", CallingConvention = CallingConvention.Cdecl)]
			public static extern Int32 mdPresortSetPSNDCZipAutomationSort(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetPSNDCZipNonAutomationSort", CallingConvention = CallingConvention.Cdecl)]
			public static extern Int32 mdPresortSetPSNDCZipNonAutomationSort(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetPSNDCZipECRRTSort", CallingConvention = CallingConvention.Cdecl)]
			public static extern Int32 mdPresortSetPSNDCZipECRRTSort(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetPSNDCZipCoSack", CallingConvention = CallingConvention.Cdecl)]
			public static extern Int32 mdPresortSetPSNDCZipCoSack(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetPSDDUZipECRRTSort", CallingConvention = CallingConvention.Cdecl)]
			public static extern Int32 mdPresortSetPSDDUZipECRRTSort(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetPSPOMZipAutomationSort", CallingConvention = CallingConvention.Cdecl)]
			public static extern Int32 mdPresortSetPSPOMZipAutomationSort(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetPSPOMZipNonAutomationSort", CallingConvention = CallingConvention.Cdecl)]
			public static extern Int32 mdPresortSetPSPOMZipNonAutomationSort(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetPSPOMZipECRRTSort", CallingConvention = CallingConvention.Cdecl)]
			public static extern Int32 mdPresortSetPSPOMZipECRRTSort(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetPSPOMZipCoSack", CallingConvention = CallingConvention.Cdecl)]
			public static extern Int32 mdPresortSetPSPOMZipCoSack(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetPSIncludeResidualPieces", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetPSIncludeResidualPieces(IntPtr i, Int32 p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortProducePostStatementForResidualPieces", CallingConvention = CallingConvention.Cdecl)]
			public static extern Int32 mdPresortProducePostStatementForResidualPieces(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortProducePostStatement", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortProducePostStatement(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetPostStatementToSelect", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetPostStatementToSelect(IntPtr i, Int32 p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetIgnoreAspectRatio", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetIgnoreAspectRatio(IntPtr i, Int32 p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetProduceIMBCode", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetProduceIMBCode(IntPtr i, Int32 p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetIMBNumericCode", CallingConvention = CallingConvention.Cdecl)]
			public static extern IntPtr mdPresortGetIMBNumericCode(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetIMBAlphaCode", CallingConvention = CallingConvention.Cdecl)]
			public static extern IntPtr mdPresortGetIMBAlphaCode(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortGetIMBSerialNumber", CallingConvention = CallingConvention.Cdecl)]
			public static extern IntPtr mdPresortGetIMBSerialNumber(IntPtr i);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetPSMailingAgentCRID", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetPSMailingAgentCRID(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetContainerSequenceNumber", CallingConvention = CallingConvention.Cdecl)]
			public static extern Int32 mdPresortSetContainerSequenceNumber(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetPSNonProfitAuthNumberMO", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetPSNonProfitAuthNumberMO(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetProduceDropShipForms", CallingConvention = CallingConvention.Cdecl)]
			public static extern Int32 mdPresortSetProduceDropShipForms(IntPtr i, Int32 p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetExpandedReportName", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetExpandedReportName(IntPtr i, Int32 p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortProduceMailDatFiles", CallingConvention = CallingConvention.Cdecl)]
			public static extern Int32 mdPresortProduceMailDatFiles(IntPtr i, IntPtr p1, IntPtr p2);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetPSPostOfficeOfMailingPlus4", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetPSPostOfficeOfMailingPlus4(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetMDACSKeyLineData", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetMDACSKeyLineData(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetMDMachineID", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetMDMachineID(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetMDJobID", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetMDJobID(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetMDHDRIDEAllianceVersion", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetMDHDRIDEAllianceVersion(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetMDHDRHistorySequenceNumber", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetMDHDRHistorySequenceNumber(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetMDHDRHistoryStatus", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetMDHDRHistoryStatus(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetMDHDRHistoricalJobID", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetMDHDRHistoricalJobID(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetMDHDRLicensedUsersJobNumber", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetMDHDRLicensedUsersJobNumber(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetMDHDRJobNameTitleIssue", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetMDHDRJobNameTitleIssue(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetMDHDRFileSource", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetMDHDRFileSource(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetMDHDRUserLicenseCode", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetMDHDRUserLicenseCode(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetMDHDRMailDatSoftwareVendorName", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetMDHDRMailDatSoftwareVendorName(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetMDHDRMailDatSoftwareProductsName", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetMDHDRMailDatSoftwareProductsName(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetMDHDRMailDatSoftwareVersion", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetMDHDRMailDatSoftwareVersion(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetMDHDRMailDatSoftwareVendorEmail", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetMDHDRMailDatSoftwareVendorEmail(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetMDHDReDocSenderCRID", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetMDHDReDocSenderCRID(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetMDSEGVerificationFacilityName", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetMDSEGVerificationFacilityName(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetMDSEGVerificationFacilityZipPlus4", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetMDSEGVerificationFacilityZipPlus4(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetMDHDROriginalSoftwareVendorName", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetMDHDROriginalSoftwareVendorName(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetMDHDROriginalSoftwareProductsName", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetMDHDROriginalSoftwareProductsName(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetMDHDROriginalSoftwareVersion", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetMDHDROriginalSoftwareVersion(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetMDHDROriginalSoftwareVendorEmail", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetMDHDROriginalSoftwareVendorEmail(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetMDHDRContactName", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetMDHDRContactName(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetMDHDRContactPhone", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetMDHDRContactPhone(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetMDHDRContactEMail", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetMDHDRContactEMail(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetMDCPTMailOwnerID", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetMDCPTMailOwnerID(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetMDCPTOwnerCRID", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetMDCPTOwnerCRID(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetMDCPTMailOwnersMailingRefID", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetMDCPTMailOwnersMailingRefID(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetMDCPTPostalPriceIncID", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetMDCPTPostalPriceIncID(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetMDCPTPostalPriceIncType", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetMDCPTPostalPriceIncType(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetMDCPTStandParcelType", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetMDCPTStandParcelType(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetMDCPTStandFlatType", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetMDCPTStandFlatType(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetMDCPTUserOptField", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetMDCPTUserOptField(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetMDCPTContentOfMail", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetMDCPTContentOfMail(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetMDCSMCSAAgreementID", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetMDCSMCSAAgreementID(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetMDSEGDescription", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetMDSEGDescription(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetMDCPTComDescription", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetMDCPTComDescription(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetMDMPUName", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetMDMPUName(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetMDMPUDescription", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetMDMPUDescription(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetMDMPADescription", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetMDMPADescription(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetMDMPAMailingAgentMailerID", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetMDMPAMailingAgentMailerID(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetMDMPAMailOwnerPermitNumber", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetMDMPAMailOwnerPermitNumber(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetMDMPAMailOwnerPermitType", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetMDMPAMailOwnerPermitType(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetMDMPAAdditionalPostagePaymentMethod", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetMDMPAAdditionalPostagePaymentMethod(IntPtr i, IntPtr val);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetMDMPAAdditionalPaymentPermitNumber", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetMDMPAAdditionalPaymentPermitNumber(IntPtr i, IntPtr val);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetMDMPAAdditionalPaymentPermitZipPlus4", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetMDMPAAdditionalPaymentPermitZipPlus4(IntPtr i, IntPtr val);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetMDCCRCharacteristic", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetMDCCRCharacteristic(IntPtr i, IntPtr val);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetMDCCRCharacteristicType", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetMDCCRCharacteristicType(IntPtr i, IntPtr val);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetPOMLocaleKey", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetPOMLocaleKey(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetDDULocaleKey", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetDDULocaleKey(IntPtr i, IntPtr p1);
			[DllImport("mdPresort", EntryPoint = "mdPresortSetDDUPostalCode", CallingConvention = CallingConvention.Cdecl)]
			public static extern void mdPresortSetDDUPostalCode(IntPtr i, IntPtr p1);
		}

		public mdPresort() {
			i = mdPresortUnmanaged.mdPresortCreate();
		}

		~mdPresort() {
			Dispose();
		}

		public virtual void Dispose() {
			lock (this) {
				mdPresortUnmanaged.mdPresortDestroy(i);
				GC.SuppressFinalize(this);
			}
		}

		public ProgramStatus InitializeDataFiles() {
			return (ProgramStatus)mdPresortUnmanaged.mdPresortInitializeDataFiles(i);
		}

		public bool SetLicenseString(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			return (mdPresortUnmanaged.mdPresortSetLicenseString(i, u_p1.GetUtf8Ptr()) != 0);
		}

		public void SetPathToPresortDataFiles(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetPathToPresortDataFiles(i, u_p1.GetUtf8Ptr());
		}

		public string GetInitializeErrorString() {
			return Utf8String.GetUnicodeString(mdPresortUnmanaged.mdPresortGetInitializeErrorString(i));
		}

		public string GetParametersErrorString() {
			return Utf8String.GetUnicodeString(mdPresortUnmanaged.mdPresortGetParametersErrorString(i));
		}

		public string GetBuildNumber() {
			return Utf8String.GetUnicodeString(mdPresortUnmanaged.mdPresortGetBuildNumber(i));
		}

		public string GetLicenseStringExpirationDate() {
			return Utf8String.GetUnicodeString(mdPresortUnmanaged.mdPresortGetLicenseStringExpirationDate(i));
		}

		public string GetDatabaseExpirationDate() {
			return Utf8String.GetUnicodeString(mdPresortUnmanaged.mdPresortGetDatabaseExpirationDate(i));
		}

		public string GetDatabaseDate() {
			return Utf8String.GetUnicodeString(mdPresortUnmanaged.mdPresortGetDatabaseDate(i));
		}

		public void SetZip(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetZip(i, u_p1.GetUtf8Ptr());
		}

		public void SetPlus4(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetPlus4(i, u_p1.GetUtf8Ptr());
		}

		public void SetCarrierRoute(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetCarrierRoute(i, u_p1.GetUtf8Ptr());
		}

		public void SetWalkSequence(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetWalkSequence(i, u_p1.GetUtf8Ptr());
		}

		public void SetDeliveryPointCode(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetDeliveryPointCode(i, u_p1.GetUtf8Ptr());
		}

		public void SetBusinessResidentialIndicator(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetBusinessResidentialIndicator(i, u_p1.GetUtf8Ptr());
		}

		public void SetPSPermitImprint(bool p1) {
			mdPresortUnmanaged.mdPresortSetPSPermitImprint(i, (p1 ? 1 : 0));
		}

		public void SetLOTNumber(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetLOTNumber(i, u_p1.GetUtf8Ptr());
		}

		public void SetLOTOrder(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetLOTOrder(i, u_p1.GetUtf8Ptr());
		}

		public void SetPSASE(bool p1) {
			mdPresortUnmanaged.mdPresortSetPSASE(i, (p1 ? 1 : 0));
		}

		public void SetPSFASTforward(bool p1) {
			mdPresortUnmanaged.mdPresortSetPSFASTforward(i, (p1 ? 1 : 0));
		}

		public void SetPSNCOA(bool p1) {
			mdPresortUnmanaged.mdPresortSetPSNCOA(i, (p1 ? 1 : 0));
		}

		public void SetPSACS(bool p1) {
			mdPresortUnmanaged.mdPresortSetPSACS(i, (p1 ? 1 : 0));
		}

		public void SetPSAltMethod(bool p1) {
			mdPresortUnmanaged.mdPresortSetPSAltMethod(i, (p1 ? 1 : 0));
		}

		public void SetPSMultiple(bool p1) {
			mdPresortUnmanaged.mdPresortSetPSMultiple(i, (p1 ? 1 : 0));
		}

		public void SetPSOneCode(bool p1) {
			mdPresortUnmanaged.mdPresortSetPSOneCode(i, (p1 ? 1 : 0));
		}

		public void SetPSAltAddFmt(bool p1) {
			mdPresortUnmanaged.mdPresortSetPSAltAddFmt(i, (p1 ? 1 : 0));
		}

		public void SetPSCASSDate(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetPSCASSDate(i, u_p1.GetUtf8Ptr());
		}

		public void SetPSPrecanceledStamp(bool p1) {
			mdPresortUnmanaged.mdPresortSetPSPrecanceledStamp(i, (p1 ? 1 : 0));
		}

		public void SetPSPermitHolderName(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetPSPermitHolderName(i, u_p1.GetUtf8Ptr());
		}

		public void SetPSPermitHolderCompany(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetPSPermitHolderCompany(i, u_p1.GetUtf8Ptr());
		}

		public void SetPSPermitHolderStreet(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetPSPermitHolderStreet(i, u_p1.GetUtf8Ptr());
		}

		public void SetPSPermitHolderCity(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetPSPermitHolderCity(i, u_p1.GetUtf8Ptr());
		}

		public void SetPSPermitHolderState(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetPSPermitHolderState(i, u_p1.GetUtf8Ptr());
		}

		public void SetPSPermitHolderPhone(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetPSPermitHolderPhone(i, u_p1.GetUtf8Ptr());
		}

		public void SetPSPermitHolderEmail(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetPSPermitHolderEmail(i, u_p1.GetUtf8Ptr());
		}

		public void SetPSPermitHolderListName(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetPSPermitHolderListName(i, u_p1.GetUtf8Ptr());
		}

		public void SetPSPermitHolderZIP(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetPSPermitHolderZIP(i, u_p1.GetUtf8Ptr());
		}

		public void SetPSPostOfficeOfMailingCity(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetPSPostOfficeOfMailingCity(i, u_p1.GetUtf8Ptr());
		}

		public void SetPSPostOfficeOfMailingState(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetPSPostOfficeOfMailingState(i, u_p1.GetUtf8Ptr());
		}

		public void SetPSPostOfficeOfMailingZIP(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetPSPostOfficeOfMailingZIP(i, u_p1.GetUtf8Ptr());
		}

		public void SetPSMailingAgentName(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetPSMailingAgentName(i, u_p1.GetUtf8Ptr());
		}

		public void SetPSMailingAgentCompany(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetPSMailingAgentCompany(i, u_p1.GetUtf8Ptr());
		}

		public void SetPSMailingAgentStreet(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetPSMailingAgentStreet(i, u_p1.GetUtf8Ptr());
		}

		public void SetPSMailingAgentCity(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetPSMailingAgentCity(i, u_p1.GetUtf8Ptr());
		}

		public void SetPSMailingAgentState(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetPSMailingAgentState(i, u_p1.GetUtf8Ptr());
		}

		public void SetPSMailingAgentPhone(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetPSMailingAgentPhone(i, u_p1.GetUtf8Ptr());
		}

		public void SetPSMailingAgentZIP(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetPSMailingAgentZIP(i, u_p1.GetUtf8Ptr());
		}

		public void SetPSIndividualOrOrganizationName(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetPSIndividualOrOrganizationName(i, u_p1.GetUtf8Ptr());
		}

		public void SetPSIndividualOrOrganizationCompany(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetPSIndividualOrOrganizationCompany(i, u_p1.GetUtf8Ptr());
		}

		public void SetPSIndividualOrOrganizationStreet(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetPSIndividualOrOrganizationStreet(i, u_p1.GetUtf8Ptr());
		}

		public void SetPSIndividualOrOrganizationCity(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetPSIndividualOrOrganizationCity(i, u_p1.GetUtf8Ptr());
		}

		public void SetPSIndividualOrOrganizationState(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetPSIndividualOrOrganizationState(i, u_p1.GetUtf8Ptr());
		}

		public void SetPSIndividualOrOrganizationZIP(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetPSIndividualOrOrganizationZIP(i, u_p1.GetUtf8Ptr());
		}

		public void SetPSIndividualOrOrganizationCRID(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetPSIndividualOrOrganizationCRID(i, u_p1.GetUtf8Ptr());
		}

		public void SetContinueContainerNumber(bool p1) {
			mdPresortUnmanaged.mdPresortSetContinueContainerNumber(i, (p1 ? 1 : 0));
		}

		public bool SetPOMasNDC(bool p1) {
			return (mdPresortUnmanaged.mdPresortSetPOMasNDC(i, (p1 ? 1 : 0)) != 0);
		}

		public bool SetPOMasSCF(bool p1) {
			return (mdPresortUnmanaged.mdPresortSetPOMasSCF(i, (p1 ? 1 : 0)) != 0);
		}

		public void SetPSStatementSeqNumber(string val) {
			Utf8String u_val = new Utf8String(val);
			mdPresortUnmanaged.mdPresortSetPSStatementSeqNumber(i, u_val.GetUtf8Ptr());
		}

		public void SetPSMailingDate(string val) {
			Utf8String u_val = new Utf8String(val);
			mdPresortUnmanaged.mdPresortSetPSMailingDate(i, u_val.GetUtf8Ptr());
		}

		public void SetPSFedAgencyCode(string val) {
			Utf8String u_val = new Utf8String(val);
			mdPresortUnmanaged.mdPresortSetPSFedAgencyCode(i, u_val.GetUtf8Ptr());
		}

		public void SetPSCustomerNumber(string val) {
			Utf8String u_val = new Utf8String(val);
			mdPresortUnmanaged.mdPresortSetPSCustomerNumber(i, u_val.GetUtf8Ptr());
		}

		public void SetPSCAPSNumber(string val) {
			Utf8String u_val = new Utf8String(val);
			mdPresortUnmanaged.mdPresortSetPSCAPSNumber(i, u_val.GetUtf8Ptr());
		}

		public void SetPermitNumber(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetPermitNumber(i, u_p1.GetUtf8Ptr());
		}

		public void SetPSNonProfitAuthNumber(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetPSNonProfitAuthNumber(i, u_p1.GetUtf8Ptr());
		}

		public void SetMailClass(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetMailClass(i, u_p1.GetUtf8Ptr());
		}

		public void SetPieceType(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetPieceType(i, u_p1.GetUtf8Ptr());
		}

		public void SetSortType(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetSortType(i, u_p1.GetUtf8Ptr());
		}

		public void SetFCM_CRRT(bool p1) {
			mdPresortUnmanaged.mdPresortSetFCM_CRRT(i, (p1 ? 1 : 0));
		}

		public void SetFCM_CRRT_5(bool p1) {
			mdPresortUnmanaged.mdPresortSetFCM_CRRT_5(i, (p1 ? 1 : 0));
		}

		public void SetFCM_CRRT_3(bool p1) {
			mdPresortUnmanaged.mdPresortSetFCM_CRRT_3(i, (p1 ? 1 : 0));
		}

		public void SetFCM_5dg_Scheme(bool p1) {
			mdPresortUnmanaged.mdPresortSetFCM_5dg_Scheme(i, (p1 ? 1 : 0));
		}

		public void SetFCM_Auto_5dg(bool p1) {
			mdPresortUnmanaged.mdPresortSetFCM_Auto_5dg(i, (p1 ? 1 : 0));
		}

		public void SetFCM_NonAuto_5dg(bool p1) {
			mdPresortUnmanaged.mdPresortSetFCM_NonAuto_5dg(i, (p1 ? 1 : 0));
		}

		public void SetSTD_Auto_5dg_Scheme(bool p1) {
			mdPresortUnmanaged.mdPresortSetSTD_Auto_5dg_Scheme(i, (p1 ? 1 : 0));
		}

		public void SetSTD_Auto_5dg(bool p1) {
			mdPresortUnmanaged.mdPresortSetSTD_Auto_5dg(i, (p1 ? 1 : 0));
		}

		public void SetNonMachinableType(bool p1) {
			mdPresortUnmanaged.mdPresortSetNonMachinableType(i, (p1 ? 1 : 0));
		}

		public void SetRecordID(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetRecordID(i, u_p1.GetUtf8Ptr());
		}

		public void SetSackWeight(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetSackWeight(i, u_p1.GetUtf8Ptr());
		}

		public void SetPieceLength(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetPieceLength(i, u_p1.GetUtf8Ptr());
		}

		public void SetPieceHeight(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetPieceHeight(i, u_p1.GetUtf8Ptr());
		}

		public void SetPieceThickness(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetPieceThickness(i, u_p1.GetUtf8Ptr());
		}

		public void SetPieceWeight(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetPieceWeight(i, u_p1.GetUtf8Ptr());
		}

		public void SetPackageSize(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetPackageSize(i, u_p1.GetUtf8Ptr());
		}

		public void SetEnableCoTray(bool p1) {
			mdPresortUnmanaged.mdPresortSetEnableCoTray(i, (p1 ? 1 : 0));
		}

		public void SetUseFSM1000(bool p1) {
			mdPresortUnmanaged.mdPresortSetUseFSM1000(i, (p1 ? 1 : 0));
		}

		public void SetSTDNonProfit(bool p1) {
			mdPresortUnmanaged.mdPresortSetSTDNonProfit(i, (p1 ? 1 : 0));
		}

		public void SetIgnoreDSF(bool val) {
			mdPresortUnmanaged.mdPresortSetIgnoreDSF(i, (val ? 1 : 0));
		}

		public void SetMailersID(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetMailersID(i, u_p1.GetUtf8Ptr());
		}

		public bool SetPSPrecanceledStampValue(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			return (mdPresortUnmanaged.mdPresortSetPSPrecanceledStampValue(i, u_p1.GetUtf8Ptr()) != 0);
		}

		public void SetPSPresortResidualPieces(bool p1) {
			mdPresortUnmanaged.mdPresortSetPSPresortResidualPieces(i, (p1 ? 1 : 0));
		}

		public void SetIMBSerialNumber(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetIMBSerialNumber(i, u_p1.GetUtf8Ptr());
		}

		public void SetPSPolicitalCampaignMailing(bool p1) {
			mdPresortUnmanaged.mdPresortSetPSPolicitalCampaignMailing(i, (p1 ? 1 : 0));
		}

		public void SetPSOfficialElectionMail(bool p1) {
			mdPresortUnmanaged.mdPresortSetPSOfficialElectionMail(i, (p1 ? 1 : 0));
		}

		public bool ProduceReports(string p1, string p2) {
			Utf8String u_p1 = new Utf8String(p1);
			Utf8String u_p2 = new Utf8String(p2);
			return (mdPresortUnmanaged.mdPresortProduceReports(i, u_p1.GetUtf8Ptr(), u_p2.GetUtf8Ptr()) != 0);
		}

		public void SetSCFCity(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetSCFCity(i, u_p1.GetUtf8Ptr());
		}

		public void SetSCFState(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetSCFState(i, u_p1.GetUtf8Ptr());
		}

		public void SetSCFZip(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetSCFZip(i, u_p1.GetUtf8Ptr());
		}

		public bool AddSCF() {
			return (mdPresortUnmanaged.mdPresortAddSCF(i) != 0);
		}

		public void SetNDCCity(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetNDCCity(i, u_p1.GetUtf8Ptr());
		}

		public void SetNDCState(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetNDCState(i, u_p1.GetUtf8Ptr());
		}

		public void SetNDCZip(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetNDCZip(i, u_p1.GetUtf8Ptr());
		}

		public bool AddNDC() {
			return (mdPresortUnmanaged.mdPresortAddNDC(i) != 0);
		}

		public void SetDDUCity(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetDDUCity(i, u_p1.GetUtf8Ptr());
		}

		public void SetDDUState(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetDDUState(i, u_p1.GetUtf8Ptr());
		}

		public void SetDDUZip(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetDDUZip(i, u_p1.GetUtf8Ptr());
		}

		public void SetDDUMoreZip(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetDDUMoreZip(i, u_p1.GetUtf8Ptr());
		}

		public bool AddDDU() {
			return (mdPresortUnmanaged.mdPresortAddDDU(i) != 0);
		}

		public void SetTTBorder(bool p1) {
			mdPresortUnmanaged.mdPresortSetTTBorder(i, (p1 ? 1 : 0));
		}

		public void SetTTNumberofPieces(bool p1) {
			mdPresortUnmanaged.mdPresortSetTTNumberofPieces(i, (p1 ? 1 : 0));
		}

		public void SetTTContainerNumber(bool val) {
			mdPresortUnmanaged.mdPresortSetTTContainerNumber(i, (val ? 1 : 0));
		}

		public void SetTTContainerSize(bool val) {
			mdPresortUnmanaged.mdPresortSetTTContainerSize(i, (val ? 1 : 0));
		}

		public void SetTTOther(string val2) {
			Utf8String u_val2 = new Utf8String(val2);
			mdPresortUnmanaged.mdPresortSetTTOther(i, u_val2.GetUtf8Ptr());
		}

		public void SetTTParameterPositionX(string val) {
			Utf8String u_val = new Utf8String(val);
			mdPresortUnmanaged.mdPresortSetTTParameterPositionX(i, u_val.GetUtf8Ptr());
		}

		public void SetTTParameterPositionY(string val) {
			Utf8String u_val = new Utf8String(val);
			mdPresortUnmanaged.mdPresortSetTTParameterPositionY(i, u_val.GetUtf8Ptr());
		}

		public void SetTTParameterWidth(string val) {
			Utf8String u_val = new Utf8String(val);
			mdPresortUnmanaged.mdPresortSetTTParameterWidth(i, u_val.GetUtf8Ptr());
		}

		public void SetTTParameterHeight(string val) {
			Utf8String u_val = new Utf8String(val);
			mdPresortUnmanaged.mdPresortSetTTParameterHeight(i, u_val.GetUtf8Ptr());
		}

		public bool DoPresort() {
			return (mdPresortUnmanaged.mdPresortDoPresort(i) != 0);
		}

		public void SetDisableDetailedOutput(bool p1) {
			mdPresortUnmanaged.mdPresortSetDisableDetailedOutput(i, (p1 ? 1 : 0));
		}

		public string SetPreSortSettings(int p1) {
			return Utf8String.GetUnicodeString(mdPresortUnmanaged.mdPresortSetPreSortSettings(i, p1));
		}

		public bool SetACSCodeSettings(int p1) {
			return (mdPresortUnmanaged.mdPresortSetACSCodeSettings(i, p1) != 0);
		}

		public bool UpdateParameters() {
			return (mdPresortUnmanaged.mdPresortUpdateParameters(i) != 0);
		}

		public bool AddRecord() {
			return (mdPresortUnmanaged.mdPresortAddRecord(i) != 0);
		}

		public bool GetRecord(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			return (mdPresortUnmanaged.mdPresortGetRecord(i, u_p1.GetUtf8Ptr()) != 0);
		}

		public bool GetFirstRecord() {
			return (mdPresortUnmanaged.mdPresortGetFirstRecord(i) != 0);
		}

		public bool GetNextRecord() {
			return (mdPresortUnmanaged.mdPresortGetNextRecord(i) != 0);
		}

		public string GetMailPieceRate() {
			return Utf8String.GetUnicodeString(mdPresortUnmanaged.mdPresortGetMailPieceRate(i));
		}

		public string GetZipAsString() {
			return Utf8String.GetUnicodeString(mdPresortUnmanaged.mdPresortGetZipAsString(i));
		}

		public int GetTrayNumber() {
			return mdPresortUnmanaged.mdPresortGetTrayNumber(i);
		}

		public int GetSequenceNumber() {
			return mdPresortUnmanaged.mdPresortGetSequenceNumber(i);
		}

		public string GetRecordID() {
			return Utf8String.GetUnicodeString(mdPresortUnmanaged.mdPresortGetRecordID(i));
		}

		public string GetEndorsementLine() {
			return Utf8String.GetUnicodeString(mdPresortUnmanaged.mdPresortGetEndorsementLine(i));
		}

		public string GetRateCode() {
			return Utf8String.GetUnicodeString(mdPresortUnmanaged.mdPresortGetRateCode(i));
		}

		public string GetMailJob() {
			return Utf8String.GetUnicodeString(mdPresortUnmanaged.mdPresortGetMailJob(i));
		}

		public string GetSortLevel() {
			return Utf8String.GetUnicodeString(mdPresortUnmanaged.mdPresortGetSortLevel(i));
		}

		public string GetLabelLine1() {
			return Utf8String.GetUnicodeString(mdPresortUnmanaged.mdPresortGetLabelLine1(i));
		}

		public string GetLabelLine2() {
			return Utf8String.GetUnicodeString(mdPresortUnmanaged.mdPresortGetLabelLine2(i));
		}

		public string GetLabelList() {
			return Utf8String.GetUnicodeString(mdPresortUnmanaged.mdPresortGetLabelList(i));
		}

		public string GetCINCode() {
			return Utf8String.GetUnicodeString(mdPresortUnmanaged.mdPresortGetCINCode(i));
		}

		public string GetTrayProcessingCode() {
			return Utf8String.GetUnicodeString(mdPresortUnmanaged.mdPresortGetTrayProcessingCode(i));
		}

		public string GetMailSplitCode() {
			return Utf8String.GetUnicodeString(mdPresortUnmanaged.mdPresortGetMailSplitCode(i));
		}

		public string GetTrayZipCode() {
			return Utf8String.GetUnicodeString(mdPresortUnmanaged.mdPresortGetTrayZipCode(i));
		}

		public string GetTrayType() {
			return Utf8String.GetUnicodeString(mdPresortUnmanaged.mdPresortGetTrayType(i));
		}

		public string GetBundleZipCode() {
			return Utf8String.GetUnicodeString(mdPresortUnmanaged.mdPresortGetBundleZipCode(i));
		}

		public string GetBundleSortLevel() {
			return Utf8String.GetUnicodeString(mdPresortUnmanaged.mdPresortGetBundleSortLevel(i));
		}

		public int GetBundleNumber() {
			return mdPresortUnmanaged.mdPresortGetBundleNumber(i);
		}

		public string GetBundleIndicator() {
			return Utf8String.GetUnicodeString(mdPresortUnmanaged.mdPresortGetBundleIndicator(i));
		}

		public string GetZipCodeInScheme() {
			return Utf8String.GetUnicodeString(mdPresortUnmanaged.mdPresortGetZipCodeInScheme(i));
		}

		public string GetBarcodeID() {
			return Utf8String.GetUnicodeString(mdPresortUnmanaged.mdPresortGetBarcodeID(i));
		}

		public string GetServiceTypeID() {
			return Utf8String.GetUnicodeString(mdPresortUnmanaged.mdPresortGetServiceTypeID(i));
		}

		public bool SetProducePallets(bool p1) {
			return (mdPresortUnmanaged.mdPresortSetProducePallets(i, (p1 ? 1 : 0)) != 0);
		}

		public void SetIgnoreMinSCFPallet(bool val) {
			mdPresortUnmanaged.mdPresortSetIgnoreMinSCFPallet(i, (val ? 1 : 0));
		}

		public int GetPalletsTotal() {
			return mdPresortUnmanaged.mdPresortGetPalletsTotal(i);
		}

		public string GetPalletZipCode() {
			return Utf8String.GetUnicodeString(mdPresortUnmanaged.mdPresortGetPalletZipCode(i));
		}

		public string GetPalletLabelLine1() {
			return Utf8String.GetUnicodeString(mdPresortUnmanaged.mdPresortGetPalletLabelLine1(i));
		}

		public string GetPalletLabelLine2() {
			return Utf8String.GetUnicodeString(mdPresortUnmanaged.mdPresortGetPalletLabelLine2(i));
		}

		public string GetPalletSortLevel() {
			return Utf8String.GetUnicodeString(mdPresortUnmanaged.mdPresortGetPalletSortLevel(i));
		}

		public int GetPalletNumber() {
			return mdPresortUnmanaged.mdPresortGetPalletNumber(i);
		}

		public int GetTotalPiecesAuto5dig() {
			return mdPresortUnmanaged.mdPresortGetTotalPiecesAuto5dig(i);
		}

		public int GetTotalPiecesAuto3dig() {
			return mdPresortUnmanaged.mdPresortGetTotalPiecesAuto3dig(i);
		}

		public int GetTotalPiecesAutoADC() {
			return mdPresortUnmanaged.mdPresortGetTotalPiecesAutoADC(i);
		}

		public int GetTotalPiecesAutoMADC() {
			return mdPresortUnmanaged.mdPresortGetTotalPiecesAutoMADC(i);
		}

		public int GetTotalPieces5dig() {
			return mdPresortUnmanaged.mdPresortGetTotalPieces5dig(i);
		}

		public int GetTotalPieces3dig() {
			return mdPresortUnmanaged.mdPresortGetTotalPieces3dig(i);
		}

		public int GetTotalPiecesADC() {
			return mdPresortUnmanaged.mdPresortGetTotalPiecesADC(i);
		}

		public int GetTotalPiecesMADC() {
			return mdPresortUnmanaged.mdPresortGetTotalPiecesMADC(i);
		}

		public int GetTotalPiecesWS() {
			return mdPresortUnmanaged.mdPresortGetTotalPiecesWS(i);
		}

		public int GetTotalPiecesHDP() {
			return mdPresortUnmanaged.mdPresortGetTotalPiecesHDP(i);
		}

		public int GetTotalPiecesHD() {
			return mdPresortUnmanaged.mdPresortGetTotalPiecesHD(i);
		}

		public int GetTotalPiecesLOT() {
			return mdPresortUnmanaged.mdPresortGetTotalPiecesLOT(i);
		}

		public double GetRateAuto5dig() {
			return mdPresortUnmanaged.mdPresortGetRateAuto5dig(i);
		}

		public double GetRateAuto3dig() {
			return mdPresortUnmanaged.mdPresortGetRateAuto3dig(i);
		}

		public double GetRateAutoADC() {
			return mdPresortUnmanaged.mdPresortGetRateAutoADC(i);
		}

		public double GetRateAutoMADC() {
			return mdPresortUnmanaged.mdPresortGetRateAutoMADC(i);
		}

		public double GetRate5dig() {
			return mdPresortUnmanaged.mdPresortGetRate5dig(i);
		}

		public double GetRate3dig() {
			return mdPresortUnmanaged.mdPresortGetRate3dig(i);
		}

		public double GetRateADC() {
			return mdPresortUnmanaged.mdPresortGetRateADC(i);
		}

		public double GetRateMADC() {
			return mdPresortUnmanaged.mdPresortGetRateMADC(i);
		}

		public double GetRateWS() {
			return mdPresortUnmanaged.mdPresortGetRateWS(i);
		}

		public double GetRateHDP() {
			return mdPresortUnmanaged.mdPresortGetRateHDP(i);
		}

		public double GetRateHD() {
			return mdPresortUnmanaged.mdPresortGetRateHD(i);
		}

		public double GetRateLOT() {
			return mdPresortUnmanaged.mdPresortGetRateLOT(i);
		}

		public double GetPoundRate() {
			return mdPresortUnmanaged.mdPresortGetPoundRate(i);
		}

		public double GetPoundRateWS() {
			return mdPresortUnmanaged.mdPresortGetPoundRateWS(i);
		}

		public double GetPoundRateHDP() {
			return mdPresortUnmanaged.mdPresortGetPoundRateHDP(i);
		}

		public double GetPoundRateHD() {
			return mdPresortUnmanaged.mdPresortGetPoundRateHD(i);
		}

		public double GetPoundRateLOT() {
			return mdPresortUnmanaged.mdPresortGetPoundRateLOT(i);
		}

		public int GetTotalTrays1Ft() {
			return mdPresortUnmanaged.mdPresortGetTotalTrays1Ft(i);
		}

		public int GetTotalTrays2Ft() {
			return mdPresortUnmanaged.mdPresortGetTotalTrays2Ft(i);
		}

		public int GetTotalSacks() {
			return mdPresortUnmanaged.mdPresortGetTotalSacks(i);
		}

		public int GetTotalResidualPieces() {
			return mdPresortUnmanaged.mdPresortGetTotalResidualPieces(i);
		}

		public int GetTotalPiecesAuto5digFSS() {
			return mdPresortUnmanaged.mdPresortGetTotalPiecesAuto5digFSS(i);
		}

		public int GetTotalPieces5digFSS() {
			return mdPresortUnmanaged.mdPresortGetTotalPieces5digFSS(i);
		}

		public double GetRateAuto5digFSS() {
			return mdPresortUnmanaged.mdPresortGetRateAuto5digFSS(i);
		}

		public double GetRate5digFSS() {
			return mdPresortUnmanaged.mdPresortGetRate5digFSS(i);
		}

		public double GetPoundRateAuto5digFSS() {
			return mdPresortUnmanaged.mdPresortGetPoundRateAuto5digFSS(i);
		}

		public double GetPoundRate5digFSS() {
			return mdPresortUnmanaged.mdPresortGetPoundRate5digFSS(i);
		}

		public string GetDropShipZipPlus4() {
			return Utf8String.GetUnicodeString(mdPresortUnmanaged.mdPresortGetDropShipZipPlus4(i));
		}

		public string Get(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			return Utf8String.GetUnicodeString(mdPresortUnmanaged.mdPresortGet(i, u_p1.GetUtf8Ptr()));
		}

		public int GetTotalNonAutoPieces() {
			return mdPresortUnmanaged.mdPresortGetTotalNonAutoPieces(i);
		}

		public double GetPresortedRate() {
			return mdPresortUnmanaged.mdPresortGetPresortedRate(i);
		}

		public double GetResidualRate() {
			return mdPresortUnmanaged.mdPresortGetResidualRate(i);
		}

		public int GetDestTotalPiecesAuto5dig() {
			return mdPresortUnmanaged.mdPresortGetDestTotalPiecesAuto5dig(i);
		}

		public int GetDestTotalPiecesAuto3dig() {
			return mdPresortUnmanaged.mdPresortGetDestTotalPiecesAuto3dig(i);
		}

		public int GetDestTotalPiecesAutoADC() {
			return mdPresortUnmanaged.mdPresortGetDestTotalPiecesAutoADC(i);
		}

		public int GetDestTotalPiecesAutoMADC() {
			return mdPresortUnmanaged.mdPresortGetDestTotalPiecesAutoMADC(i);
		}

		public int GetDestTotalPieces5dig() {
			return mdPresortUnmanaged.mdPresortGetDestTotalPieces5dig(i);
		}

		public int GetDestTotalPieces3dig() {
			return mdPresortUnmanaged.mdPresortGetDestTotalPieces3dig(i);
		}

		public int GetDestTotalPiecesADC() {
			return mdPresortUnmanaged.mdPresortGetDestTotalPiecesADC(i);
		}

		public int GetDestTotalPiecesMADC() {
			return mdPresortUnmanaged.mdPresortGetDestTotalPiecesMADC(i);
		}

		public int GetDestTotalPiecesWS() {
			return mdPresortUnmanaged.mdPresortGetDestTotalPiecesWS(i);
		}

		public int GetDestTotalPiecesHDP() {
			return mdPresortUnmanaged.mdPresortGetDestTotalPiecesHDP(i);
		}

		public int GetDestTotalPiecesHD() {
			return mdPresortUnmanaged.mdPresortGetDestTotalPiecesHD(i);
		}

		public int GetDestTotalPiecesLOT() {
			return mdPresortUnmanaged.mdPresortGetDestTotalPiecesLOT(i);
		}

		public double GetDestRateAuto5dig() {
			return mdPresortUnmanaged.mdPresortGetDestRateAuto5dig(i);
		}

		public double GetDestRateAuto3dig() {
			return mdPresortUnmanaged.mdPresortGetDestRateAuto3dig(i);
		}

		public double GetDestRateAutoADC() {
			return mdPresortUnmanaged.mdPresortGetDestRateAutoADC(i);
		}

		public double GetDestRateAutoMADC() {
			return mdPresortUnmanaged.mdPresortGetDestRateAutoMADC(i);
		}

		public double GetDestRate5dig() {
			return mdPresortUnmanaged.mdPresortGetDestRate5dig(i);
		}

		public double GetDestRate3dig() {
			return mdPresortUnmanaged.mdPresortGetDestRate3dig(i);
		}

		public double GetDestRateADC() {
			return mdPresortUnmanaged.mdPresortGetDestRateADC(i);
		}

		public double GetDestRateMADC() {
			return mdPresortUnmanaged.mdPresortGetDestRateMADC(i);
		}

		public double GetDestRateWS() {
			return mdPresortUnmanaged.mdPresortGetDestRateWS(i);
		}

		public double GetDestRateHDP() {
			return mdPresortUnmanaged.mdPresortGetDestRateHDP(i);
		}

		public double GetDestRateHD() {
			return mdPresortUnmanaged.mdPresortGetDestRateHD(i);
		}

		public double GetDestRateLOT() {
			return mdPresortUnmanaged.mdPresortGetDestRateLOT(i);
		}

		public int GetContainersTotal(int p1) {
			return mdPresortUnmanaged.mdPresortGetContainersTotal(i, p1);
		}

		public int GetBundlesTotal(int p1) {
			return mdPresortUnmanaged.mdPresortGetBundlesTotal(i, p1);
		}

		public int GetPiecesTotal(int p1) {
			return mdPresortUnmanaged.mdPresortGetPiecesTotal(i, p1);
		}

		public double GetChargeTotal(int p1) {
			return mdPresortUnmanaged.mdPresortGetChargeTotal(i, p1);
		}

		public bool GetDestSCFInfo(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			return (mdPresortUnmanaged.mdPresortGetDestSCFInfo(i, u_p1.GetUtf8Ptr()) != 0);
		}

		public bool GetDestNDCInfo(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			return (mdPresortUnmanaged.mdPresortGetDestNDCInfo(i, u_p1.GetUtf8Ptr()) != 0);
		}

		public bool GetDestDDUInfo(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			return (mdPresortUnmanaged.mdPresortGetDestDDUInfo(i, u_p1.GetUtf8Ptr()) != 0);
		}

		public int GetDestContainersTotal(int p1) {
			return mdPresortUnmanaged.mdPresortGetDestContainersTotal(i, p1);
		}

		public int GetDestBundlesTotal(int p1) {
			return mdPresortUnmanaged.mdPresortGetDestBundlesTotal(i, p1);
		}

		public int GetDestPiecesTotal(int p1) {
			return mdPresortUnmanaged.mdPresortGetDestPiecesTotal(i, p1);
		}

		public double GetDestChargeTotal(int p1) {
			return mdPresortUnmanaged.mdPresortGetDestChargeTotal(i, p1);
		}

		public void SetTwoFtTrayMaximum(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetTwoFtTrayMaximum(i, u_p1.GetUtf8Ptr());
		}

		public void SetTwoFtTrayMinimum(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetTwoFtTrayMinimum(i, u_p1.GetUtf8Ptr());
		}

		public void SetOneFtTrayMaximum(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetOneFtTrayMaximum(i, u_p1.GetUtf8Ptr());
		}

		public void SetOneFtTrayMinimum(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetOneFtTrayMinimum(i, u_p1.GetUtf8Ptr());
		}

		public void SetBundleMaximum(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetBundleMaximum(i, u_p1.GetUtf8Ptr());
		}

		public bool SetPSSCFZipAutomationSort(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			return (mdPresortUnmanaged.mdPresortSetPSSCFZipAutomationSort(i, u_p1.GetUtf8Ptr()) != 0);
		}

		public bool SetPSSCFZipNonAutomationSort(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			return (mdPresortUnmanaged.mdPresortSetPSSCFZipNonAutomationSort(i, u_p1.GetUtf8Ptr()) != 0);
		}

		public bool SetPSSCFZipECRRTSort(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			return (mdPresortUnmanaged.mdPresortSetPSSCFZipECRRTSort(i, u_p1.GetUtf8Ptr()) != 0);
		}

		public bool SetPSSCFZipCoSack(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			return (mdPresortUnmanaged.mdPresortSetPSSCFZipCoSack(i, u_p1.GetUtf8Ptr()) != 0);
		}

		public bool SetPSNDCZipAutomationSort(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			return (mdPresortUnmanaged.mdPresortSetPSNDCZipAutomationSort(i, u_p1.GetUtf8Ptr()) != 0);
		}

		public bool SetPSNDCZipNonAutomationSort(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			return (mdPresortUnmanaged.mdPresortSetPSNDCZipNonAutomationSort(i, u_p1.GetUtf8Ptr()) != 0);
		}

		public bool SetPSNDCZipECRRTSort(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			return (mdPresortUnmanaged.mdPresortSetPSNDCZipECRRTSort(i, u_p1.GetUtf8Ptr()) != 0);
		}

		public bool SetPSNDCZipCoSack(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			return (mdPresortUnmanaged.mdPresortSetPSNDCZipCoSack(i, u_p1.GetUtf8Ptr()) != 0);
		}

		public bool SetPSDDUZipECRRTSort(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			return (mdPresortUnmanaged.mdPresortSetPSDDUZipECRRTSort(i, u_p1.GetUtf8Ptr()) != 0);
		}

		public bool SetPSPOMZipAutomationSort(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			return (mdPresortUnmanaged.mdPresortSetPSPOMZipAutomationSort(i, u_p1.GetUtf8Ptr()) != 0);
		}

		public bool SetPSPOMZipNonAutomationSort(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			return (mdPresortUnmanaged.mdPresortSetPSPOMZipNonAutomationSort(i, u_p1.GetUtf8Ptr()) != 0);
		}

		public bool SetPSPOMZipECRRTSort(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			return (mdPresortUnmanaged.mdPresortSetPSPOMZipECRRTSort(i, u_p1.GetUtf8Ptr()) != 0);
		}

		public bool SetPSPOMZipCoSack(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			return (mdPresortUnmanaged.mdPresortSetPSPOMZipCoSack(i, u_p1.GetUtf8Ptr()) != 0);
		}

		public void SetPSIncludeResidualPieces(bool p1) {
			mdPresortUnmanaged.mdPresortSetPSIncludeResidualPieces(i, (p1 ? 1 : 0));
		}

		public bool ProducePostStatementForResidualPieces() {
			return (mdPresortUnmanaged.mdPresortProducePostStatementForResidualPieces(i) != 0);
		}

		public void ProducePostStatement() {
			mdPresortUnmanaged.mdPresortProducePostStatement(i);
		}

		public void SetPostStatementToSelect(bool p1) {
			mdPresortUnmanaged.mdPresortSetPostStatementToSelect(i, (p1 ? 1 : 0));
		}

		public void SetIgnoreAspectRatio(bool p1) {
			mdPresortUnmanaged.mdPresortSetIgnoreAspectRatio(i, (p1 ? 1 : 0));
		}

		public void SetProduceIMBCode(bool p1) {
			mdPresortUnmanaged.mdPresortSetProduceIMBCode(i, (p1 ? 1 : 0));
		}

		public string GetIMBNumericCode() {
			return Utf8String.GetUnicodeString(mdPresortUnmanaged.mdPresortGetIMBNumericCode(i));
		}

		public string GetIMBAlphaCode() {
			return Utf8String.GetUnicodeString(mdPresortUnmanaged.mdPresortGetIMBAlphaCode(i));
		}

		public string GetIMBSerialNumber() {
			return Utf8String.GetUnicodeString(mdPresortUnmanaged.mdPresortGetIMBSerialNumber(i));
		}

		public void SetPSMailingAgentCRID(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetPSMailingAgentCRID(i, u_p1.GetUtf8Ptr());
		}

		public bool SetContainerSequenceNumber(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			return (mdPresortUnmanaged.mdPresortSetContainerSequenceNumber(i, u_p1.GetUtf8Ptr()) != 0);
		}

		public void SetPSNonProfitAuthNumberMO(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetPSNonProfitAuthNumberMO(i, u_p1.GetUtf8Ptr());
		}

		public bool SetProduceDropShipForms(bool p1) {
			return (mdPresortUnmanaged.mdPresortSetProduceDropShipForms(i, (p1 ? 1 : 0)) != 0);
		}

		public void SetExpandedReportName(bool p1) {
			mdPresortUnmanaged.mdPresortSetExpandedReportName(i, (p1 ? 1 : 0));
		}

		public bool ProduceMailDatFiles(string p1, string p2) {
			Utf8String u_p1 = new Utf8String(p1);
			Utf8String u_p2 = new Utf8String(p2);
			return (mdPresortUnmanaged.mdPresortProduceMailDatFiles(i, u_p1.GetUtf8Ptr(), u_p2.GetUtf8Ptr()) != 0);
		}

		public void SetPSPostOfficeOfMailingPlus4(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetPSPostOfficeOfMailingPlus4(i, u_p1.GetUtf8Ptr());
		}

		public void SetMDACSKeyLineData(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetMDACSKeyLineData(i, u_p1.GetUtf8Ptr());
		}

		public void SetMDMachineID(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetMDMachineID(i, u_p1.GetUtf8Ptr());
		}

		public void SetMDJobID(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetMDJobID(i, u_p1.GetUtf8Ptr());
		}

		public void SetMDHDRIDEAllianceVersion(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetMDHDRIDEAllianceVersion(i, u_p1.GetUtf8Ptr());
		}

		public void SetMDHDRHistorySequenceNumber(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetMDHDRHistorySequenceNumber(i, u_p1.GetUtf8Ptr());
		}

		public void SetMDHDRHistoryStatus(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetMDHDRHistoryStatus(i, u_p1.GetUtf8Ptr());
		}

		public void SetMDHDRHistoricalJobID(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetMDHDRHistoricalJobID(i, u_p1.GetUtf8Ptr());
		}

		public void SetMDHDRLicensedUsersJobNumber(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetMDHDRLicensedUsersJobNumber(i, u_p1.GetUtf8Ptr());
		}

		public void SetMDHDRJobNameTitleIssue(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetMDHDRJobNameTitleIssue(i, u_p1.GetUtf8Ptr());
		}

		public void SetMDHDRFileSource(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetMDHDRFileSource(i, u_p1.GetUtf8Ptr());
		}

		public void SetMDHDRUserLicenseCode(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetMDHDRUserLicenseCode(i, u_p1.GetUtf8Ptr());
		}

		public void SetMDHDRMailDatSoftwareVendorName(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetMDHDRMailDatSoftwareVendorName(i, u_p1.GetUtf8Ptr());
		}

		public void SetMDHDRMailDatSoftwareProductsName(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetMDHDRMailDatSoftwareProductsName(i, u_p1.GetUtf8Ptr());
		}

		public void SetMDHDRMailDatSoftwareVersion(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetMDHDRMailDatSoftwareVersion(i, u_p1.GetUtf8Ptr());
		}

		public void SetMDHDRMailDatSoftwareVendorEmail(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetMDHDRMailDatSoftwareVendorEmail(i, u_p1.GetUtf8Ptr());
		}

		public void SetMDHDReDocSenderCRID(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetMDHDReDocSenderCRID(i, u_p1.GetUtf8Ptr());
		}

		public void SetMDSEGVerificationFacilityName(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetMDSEGVerificationFacilityName(i, u_p1.GetUtf8Ptr());
		}

		public void SetMDSEGVerificationFacilityZipPlus4(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetMDSEGVerificationFacilityZipPlus4(i, u_p1.GetUtf8Ptr());
		}

		public void SetMDHDROriginalSoftwareVendorName(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetMDHDROriginalSoftwareVendorName(i, u_p1.GetUtf8Ptr());
		}

		public void SetMDHDROriginalSoftwareProductsName(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetMDHDROriginalSoftwareProductsName(i, u_p1.GetUtf8Ptr());
		}

		public void SetMDHDROriginalSoftwareVersion(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetMDHDROriginalSoftwareVersion(i, u_p1.GetUtf8Ptr());
		}

		public void SetMDHDROriginalSoftwareVendorEmail(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetMDHDROriginalSoftwareVendorEmail(i, u_p1.GetUtf8Ptr());
		}

		public void SetMDHDRContactName(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetMDHDRContactName(i, u_p1.GetUtf8Ptr());
		}

		public void SetMDHDRContactPhone(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetMDHDRContactPhone(i, u_p1.GetUtf8Ptr());
		}

		public void SetMDHDRContactEMail(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetMDHDRContactEMail(i, u_p1.GetUtf8Ptr());
		}

		public void SetMDCPTMailOwnerID(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetMDCPTMailOwnerID(i, u_p1.GetUtf8Ptr());
		}

		public void SetMDCPTOwnerCRID(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetMDCPTOwnerCRID(i, u_p1.GetUtf8Ptr());
		}

		public void SetMDCPTMailOwnersMailingRefID(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetMDCPTMailOwnersMailingRefID(i, u_p1.GetUtf8Ptr());
		}

		public void SetMDCPTPostalPriceIncID(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetMDCPTPostalPriceIncID(i, u_p1.GetUtf8Ptr());
		}

		public void SetMDCPTPostalPriceIncType(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetMDCPTPostalPriceIncType(i, u_p1.GetUtf8Ptr());
		}

		public void SetMDCPTStandParcelType(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetMDCPTStandParcelType(i, u_p1.GetUtf8Ptr());
		}

		public void SetMDCPTStandFlatType(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetMDCPTStandFlatType(i, u_p1.GetUtf8Ptr());
		}

		public void SetMDCPTUserOptField(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetMDCPTUserOptField(i, u_p1.GetUtf8Ptr());
		}

		public void SetMDCPTContentOfMail(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetMDCPTContentOfMail(i, u_p1.GetUtf8Ptr());
		}

		public void SetMDCSMCSAAgreementID(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetMDCSMCSAAgreementID(i, u_p1.GetUtf8Ptr());
		}

		public void SetMDSEGDescription(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetMDSEGDescription(i, u_p1.GetUtf8Ptr());
		}

		public void SetMDCPTComDescription(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetMDCPTComDescription(i, u_p1.GetUtf8Ptr());
		}

		public void SetMDMPUName(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetMDMPUName(i, u_p1.GetUtf8Ptr());
		}

		public void SetMDMPUDescription(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetMDMPUDescription(i, u_p1.GetUtf8Ptr());
		}

		public void SetMDMPADescription(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetMDMPADescription(i, u_p1.GetUtf8Ptr());
		}

		public void SetMDMPAMailingAgentMailerID(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetMDMPAMailingAgentMailerID(i, u_p1.GetUtf8Ptr());
		}

		public void SetMDMPAMailOwnerPermitNumber(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetMDMPAMailOwnerPermitNumber(i, u_p1.GetUtf8Ptr());
		}

		public void SetMDMPAMailOwnerPermitType(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetMDMPAMailOwnerPermitType(i, u_p1.GetUtf8Ptr());
		}

		public void SetMDMPAAdditionalPostagePaymentMethod(string val) {
			Utf8String u_val = new Utf8String(val);
			mdPresortUnmanaged.mdPresortSetMDMPAAdditionalPostagePaymentMethod(i, u_val.GetUtf8Ptr());
		}

		public void SetMDMPAAdditionalPaymentPermitNumber(string val) {
			Utf8String u_val = new Utf8String(val);
			mdPresortUnmanaged.mdPresortSetMDMPAAdditionalPaymentPermitNumber(i, u_val.GetUtf8Ptr());
		}

		public void SetMDMPAAdditionalPaymentPermitZipPlus4(string val) {
			Utf8String u_val = new Utf8String(val);
			mdPresortUnmanaged.mdPresortSetMDMPAAdditionalPaymentPermitZipPlus4(i, u_val.GetUtf8Ptr());
		}

		public void SetMDCCRCharacteristic(string val) {
			Utf8String u_val = new Utf8String(val);
			mdPresortUnmanaged.mdPresortSetMDCCRCharacteristic(i, u_val.GetUtf8Ptr());
		}

		public void SetMDCCRCharacteristicType(string val) {
			Utf8String u_val = new Utf8String(val);
			mdPresortUnmanaged.mdPresortSetMDCCRCharacteristicType(i, u_val.GetUtf8Ptr());
		}

		public void SetPOMLocaleKey(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetPOMLocaleKey(i, u_p1.GetUtf8Ptr());
		}

		public void SetDDULocaleKey(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetDDULocaleKey(i, u_p1.GetUtf8Ptr());
		}

		public void SetDDUPostalCode(string p1) {
			Utf8String u_p1 = new Utf8String(p1);
			mdPresortUnmanaged.mdPresortSetDDUPostalCode(i, u_p1.GetUtf8Ptr());
		}

		private class Utf8String : IDisposable {
			private IntPtr utf8String = IntPtr.Zero;

			public Utf8String(string str) {
				if (str == null)
					str = "";
				byte[] buffer = Encoding.UTF8.GetBytes(str);
				Array.Resize(ref buffer, buffer.Length + 1);
				buffer[buffer.Length - 1] = 0;
				utf8String = Marshal.AllocHGlobal(buffer.Length);
				Marshal.Copy(buffer, 0, utf8String, buffer.Length);
			}

			~Utf8String() {
				Dispose();
			}

			public virtual void Dispose() {
				lock (this) {
					Marshal.FreeHGlobal(utf8String);
					GC.SuppressFinalize(this);
				}
			}

			public IntPtr GetUtf8Ptr() {
				return utf8String;
			}

			public static string GetUnicodeString(IntPtr ptr) {
				if (ptr == IntPtr.Zero)
					return "";
				int len = 0;
				while (Marshal.ReadByte(ptr, len) != 0)
					len++;
				if (len == 0)
					return "";
				byte[] buffer = new byte[len];
				Marshal.Copy(ptr, buffer, 0, len);
				return Encoding.UTF8.GetString(buffer);
			}
		}
	}
}
