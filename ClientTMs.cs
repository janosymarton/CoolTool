using ResourcesAPI_TMList;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CoolTool
{
    class ClientTMs
    {
        public string client;
        public string sLang;
        public string tLang;
        public List<TMListResponse> TMlist = new List<TMListResponse>();
        private TMListResponse MasterTM = new TMListResponse();
        
        public ClientTMs(string[] id)
        {
            this.client = id[0];
            this.sLang = id[1];
            this.tLang = id[2];
        }

        public void findMasterTM()
        {
            try
            {
                foreach (TMListResponse tm in TMlist)
                {
                    if (tm.FriendlyName.ToLower().StartsWith((client + "_" + LanguageHelper.GetLanguageByMemoQCode(sLang).ISOCode + "-"
                        + LanguageHelper.GetLanguageByMemoQCode(tLang).ISOCode + "_Master").ToLower()))
                    {
                        MasterTM = tm;
                        return;
                    }
                }
            }
            catch (Exception ex)
            {
                Log.AddLog(ex.Message, false);
            }

            TMListResponse newMaster = new TMListResponse();
            newMaster.TMGuid = (new Guid()).ToString();
            newMaster.Client = client;
            newMaster.SourceLangCode = sLang;
            newMaster.TargetLangCode = tLang;
            newMaster.FriendlyName = client + "_" + LanguageHelper.GetLanguageByMemoQCode( sLang ).ISOCode + "-" 
                    + LanguageHelper.GetLanguageByMemoQCode( tLang ).ISOCode+ "_Master";

            MasterTM = newMaster;           

        }

        internal TMListResponse getMasterTM()
        {
            return MasterTM;
        }

        internal TMListResponse[] getOtherTM()
        {
            List<TMListResponse> resultList = new List<TMListResponse>();

            foreach (TMListResponse tm in TMlist)
            {
                if (tm.TMGuid != MasterTM.TMGuid)
                {
                    resultList.Add(tm);
                }
            }
            
            return resultList.ToArray();
        }
    }
}
