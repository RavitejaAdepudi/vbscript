using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace Dynatrace
{
    [Guid("86FBB5EE-753B-4453-AE6A-A37FEBF42D8F")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IADKCOM {
        void set(String value);
        int initialize(int argc, String argv);
        void uninitialize();
        int get_method_id(String method_name, String source_file_name, int source_line_no, String api, char is_artificial);
        int get_serial_no(int method_id, int entry_point);
        int enter(int method_id, int serial_no);
        void exit(int method_id, int serial_no);
        void capture_string(int serial_no, String argument);
        String get_tag_as_string();
        void set_tag_from_string(String tag);
        void link_client_purepath_by_string(bool synchronous, String tag);
        void start_server_purepath();
        void end_server_purepath();
    }

    [ComVisible(true)] // Properties and methods of this class will be visible to COM
    [Guid("F01916CF-64DD-4213-83EF-4324466C7F69")] // Unique identifier required from 'Create GUID' tool
    [ProgId("Dynatrace.ADK")] // Name used when invoking this component [e.g. CreateObject("HelloCOM.HelloCOMImplementation")]
    [ClassInterface(ClassInterfaceType.None)] // 'ClassInterfaceType.None' is only way to expose functionality through interfaces implemented explicitly by the class
    [ComSourceInterfaces(typeof(IADKCOM))] // Identifies the interface that will be exposed as COM event sources for the attributed class 
    public class ADK : IADKCOM
    {
        public void set(String value) {
            dynatrace_set(value);
        }
        [DllImport("dtadk.dll", EntryPoint = "dynatrace_set")]
        private static extern void dynatrace_set(String value);

        public int initialize(int argc, String argv)
        {
            return dynatrace_initialize(argc, argv);
        }
        [DllImport("dtadk.dll", EntryPoint = "dynatrace_initialize")]
        private static extern int dynatrace_initialize(int argc, String argv);

        public void uninitialize()
        {
            dynatrace_uninitialize();
        }
        [DllImport("dtadk.dll", EntryPoint = "dynatrace_uninitialize")]
        private static extern void dynatrace_uninitialize();

        public int get_method_id(String method_name, String source_file_name, int source_line_no, String api, char is_artificial)
        {
            return dynatrace_get_method_id(method_name, source_file_name, source_line_no, api, is_artificial);
        }
        [DllImport("dtadk.dll", EntryPoint = "dynatrace_get_method_id")]
        private static extern int dynatrace_get_method_id(String method_name, String source_file_name, int source_line_no, String api, char is_artificial);

        public int get_serial_no(int method_id, int entry_point)
        {
            return dynatrace_get_serial_no(method_id, entry_point);
        }
        [DllImport("dtadk.dll", EntryPoint = "dynatrace_get_serial_no")]
        private static extern int dynatrace_get_serial_no(int method_id, int entry_point);

        public int enter(int method_id, int serial_no)
        {
            return dynatrace_enter(method_id, serial_no);
        }
        [DllImport("dtadk.dll", EntryPoint = "dynatrace_enter")]
        private static extern int dynatrace_enter(int method_id, int serial_no);

        public void exit(int method_id, int serial_no)
        {
            dynatrace_exit(method_id, serial_no);
        }
        [DllImport("dtadk.dll", EntryPoint = "dynatrace_exit")]
        private static extern void dynatrace_exit(int method_id, int serial_no);

        public void capture_string(int serial_no, String argument)
        {
            dynatrace_capture_string(serial_no, ASCIIEncoding.ASCII.GetBytes(argument + "\0"));
        }
        [DllImport("dtadk.dll", EntryPoint = "dynatrace_capture_string")]
        private static extern void dynatrace_capture_string(int serial_no, byte[] argument);

        public String get_tag_as_string()
        {
            byte[] tag = new byte[256];
            dynatrace_get_tag_as_string(tag, tag.Length);
            int inx = Array.FindIndex(tag, 0, (x) => x == 0); //search for 0
            if (inx >= 0)
                return (System.Text.Encoding.ASCII.GetString(tag, 0, inx));
            else
                return (System.Text.Encoding.ASCII.GetString(tag));
        }
        [DllImport("dtadk.dll", EntryPoint = "dynatrace_get_tag_as_string")]
        private static extern int dynatrace_get_tag_as_string(byte[] buf, int size);

        public void set_tag_from_string(String tag)
        {
            dynatrace_set_tag_from_string(ASCIIEncoding.ASCII.GetBytes(tag + "\0"));
        }
        [DllImport("dtadk.dll", EntryPoint = "dynatrace_set_tag_from_string")]
        private static extern int dynatrace_set_tag_from_string(byte[] tag);

        public void link_client_purepath_by_string(bool synchronous, String tag)
        {
            dynatrace_link_client_purepath_by_string(synchronous, ASCIIEncoding.ASCII.GetBytes(tag + "\0"));
        }
        [DllImport("dtadk.dll", EntryPoint = "dynatrace_link_client_purepath_by_string")]
        private static extern void dynatrace_link_client_purepath_by_string(bool synchronous, byte[] tag);

        public void start_server_purepath()
        {
            dynatrace_start_server_purepath();
        }
        [DllImport("dtadk.dll", EntryPoint = "dynatrace_start_server_purepath")]
        private static extern void dynatrace_start_server_purepath();

        public void end_server_purepath()
        {
            dynatrace_end_server_purepath();
        }
        [DllImport("dtadk.dll", EntryPoint = "dynatrace_end_server_purepath")]
        private static extern void dynatrace_end_server_purepath();
    }
}
