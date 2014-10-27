using System;
using System.Text;
using System.Runtime.InteropServices;

namespace Volume
{
	 
	public class GetVol
	{
[DllImport("kernel32.dll")]
private static extern long GetVolumeInformation(string PathName, StringBuilder VolumeNameBuffer, UInt32 VolumeNameSize, ref UInt32 VolumeSerialNumber, ref UInt32 MaximumComponentLength, ref UInt32 FileSystemFlags, StringBuilder FileSystemNameBuffer, UInt32 FileSystemNameSize);
/// <summary>
/// Get Volume Serial Number as string
/// </summary>
/// <param name="strDriveLetter">Single letter (e.g., "C")</param>
/// <returns>string representation of Volume Serial Number</returns>
public string GetVolumeSerial(string strDriveLetter)
{
	uint serNum = 0;
	uint maxCompLen = 0;
	StringBuilder VolLabel = new StringBuilder(256);	// Label
	UInt32 VolFlags = new UInt32();
	StringBuilder FSName = new StringBuilder(256);	// File System Name
	strDriveLetter+=":\\";
long Ret = GetVolumeInformation(strDriveLetter, VolLabel, (UInt32)VolLabel.Capacity, ref serNum, ref maxCompLen, ref VolFlags, FSName, (UInt32)FSName.Capacity);

return Convert.ToString(serNum);
		}
	}
}
