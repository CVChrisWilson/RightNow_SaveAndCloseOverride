using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CompanyUpdateSingleton
{
    public sealed class UpdateStatusHandler
    {
        Dictionary<string, UpdateStatuses> _updateStatuses = new Dictionary<string, UpdateStatuses>();
        static object _lock = new object();
        private static UpdateStatusHandler instance;

        private UpdateStatusHandler()
        {
        }

        public static UpdateStatusHandler Instance
        {
            get
            {
                if (instance == null)
                {
                    instance = new UpdateStatusHandler();
                }
                return instance;
            }
        }

        public UpdateStatuses GetUpdateStatus(string CompanyId)
        {
            lock (_lock)
            {   
                if (!Instance._updateStatuses.ContainsKey(CompanyId))
                {
                    Instance._updateStatuses.Add(CompanyId, new UpdateStatuses());
                }
            }
            return Instance._updateStatuses[CompanyId];
        }
    }

    public class UpdateStatuses
    {

        public UpdateStatuses()
        {
        }
        public enum UpdateStatus
        {
            NotCalled = 0,
            Called = 1,
            Succeeded = 2,
            Failure = 3
        }

        public UpdateStatus updateCustomerStatus { get; set; }
        public UpdateStatus updateStudentStatus { get; set; }
        public UpdateStatus updateOxfordStatus { get; set; }
    }
}