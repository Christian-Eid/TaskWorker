using BusinessLogic.Interfaces;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using EmailSender.Interfaces;

namespace BusinessLogic.Services
{
    public class BusinessService : IBusiness
    {
        private readonly ILogger _logger;
        private readonly IEmailSender _emailSender;
        public BusinessService(ILogger logger, IEmailSender emailSender)
        {
            _logger = logger;
            _emailSender = emailSender;
        }

        //Read from DB + dependency to DB with Raw SQL 
        //Convert Data to Excel and Save File 
        //Send Via Email --Done
        public Task RunBusinessTask()
        {
            _emailSender.FillMessage();
        }
    }
}
