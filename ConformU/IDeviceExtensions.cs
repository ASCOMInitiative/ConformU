using ASCOM.Common.DeviceInterfaces;

namespace ConformU
{
    public interface IDeviceExtensions 
    {
        /// <summary>
        /// Extensions to retrieve the interface version as an object. Used to determine the type of the returned value (Int16, Int32 etc.) at runtime.
        /// </summary>
        public object InterfaceVersionObject { get; }
    }
}
