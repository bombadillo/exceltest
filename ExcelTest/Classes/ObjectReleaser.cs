using System;

namespace ExcelTest.Classes
{
    using Interfaces;

    public class ObjectReleaser : IReleaseObjects
    {
        private readonly ILog Logger;

        public ObjectReleaser(ILog logger)
        {
            Logger = logger;
        }

        public void Release(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                Logger.Error("Exception Occured while releasing object " + ex);
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
