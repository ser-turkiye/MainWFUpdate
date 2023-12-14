package ser;

import com.ser.blueline.IDocument;
import com.ser.blueline.IDocumentServer;
import com.ser.blueline.ISession;
import com.ser.blueline.IUser;
import com.ser.blueline.bpm.*;
import de.ser.doxis4.agentserver.UnifiedAgent;
import org.json.JSONObject;

import java.io.File;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Collection;
import java.util.List;
import java.util.UUID;


public class UpdateWFTask extends UnifiedAgent {

    ISession ses;
    IDocumentServer srv;
    IBpmService bpm;
    private ProcessHelper helper;
    @Override
    protected Object execute() {
        if (getEventTask() == null)
            return resultError("Null Document object");


        ses = getSes();
        srv = ses.getDocumentServer();
        bpm = getBpm();
        try {

            JSONObject scfg = Utils.getSystemConfig(ses);
            if(scfg.has("LICS.SPIRE_XLS")){
                com.spire.license.LicenseProvider.setLicenseKey(scfg.getString("LICS.SPIRE_XLS"));
            }

            this.helper = new ProcessHelper(ses);
            ITask task = getEventTask();
            if(task.getName().equals("Base task")){return resultSuccess("There is no mail 'Base task'");}

            IProcessInstance proi = task.getProcessInstance();
            (new File(Conf.MainWFUpdate.MainPath)).mkdir();


            //task.getCurrentWorkbasket().getName();
            List<String> wlst = new ArrayList<>();
            List<String> mlst = new ArrayList<>();
            /*
            Collection<ITask> tasks = proi.findTasks(TaskStatus.READY);
            for(ITask ftask : tasks){
                if(ftask.getCurrentWorkbasket() == null){continue;}
                IWorkbasket wbsk = ftask.getCurrentWorkbasket();
                if(!wbsk.getName().isEmpty()){
                    wlst.add(wbsk.getName());
                }

                String wuem = wbsk.getNotifyEMail();
                if(wuem != null && !wuem.isEmpty()){
                    mlst.add(wuem);
                }
            }
            */
            IWorkbasket wbsk = task.getCurrentWorkbasket();
            if(wbsk != null) {
                if (!wbsk.getName().isEmpty()) {
                    wlst.add(wbsk.getName());
                }
                String wuem = wbsk.getNotifyEMail();
                if (wuem != null && !wuem.isEmpty()) {
                    mlst.add(wuem);
                }
            }

            if(mlst.size() == 0){return resultSuccess("No mail address : " + (wbsk !=null ? wbsk.getFullName() : "-No Workbasket-"));}

            String taskCreation = (task.getCreationDate() == null ? "" : (new SimpleDateFormat("yyyyMMdd")).format(task.getCreationDate()));

            IDocument mainDocument = (IDocument) proi.getMainInformationObject();
            if(mainDocument == null){return resultSuccess("No-Main document");}

            String docNo = mainDocument.getDescriptorValue(Conf.Descriptors.DocNumber, String.class);
            docNo = (docNo == null ? "" : docNo);
            if(docNo.isEmpty()){return resultSuccess("Passed successfully");}

            String revNo = mainDocument.getDescriptorValue(Conf.Descriptors.Revision, String.class);
            revNo = (revNo == null ? "" : revNo);

            mainDocument.setDescriptorValue("ccmPrjDocWFProcessName", proi.getDisplayName());
            mainDocument.setDescriptorValue("ccmPrjDocWFTaskName", task.getName());
            mainDocument.setDescriptorValue("ccmPrjDocWFTaskCreation", taskCreation);
            mainDocument.setDescriptorValue("ccmPrjDocWFTaskRecipients",
                String.join(";", wlst)
            );
            mainDocument.commit();

            String uniqueId = UUID.randomUUID().toString();
            String prjn = proi.getDescriptorValue(Conf.Descriptors.ProjectNo, String.class);
            String mtpn = "UPDATE_WF_MAIL";
            IDocument mtpl = Utils.getTemplateDocument(prjn, mtpn, helper);
            if(mtpl == null){
                throw new Exception("Template-Document [ " + mtpn + " ] not found.");
            }

            JSONObject dbks = new JSONObject();
            dbks.put("DocNo", docNo);
            dbks.put("RevNo", revNo);
            dbks.put("Title", mainDocument.getDisplayName());
            dbks.put("Task", task.getName());


            JSONObject mcfg = Utils.getMailConfig(ses, srv, mtpn);
            dbks.put("DoxisLink", mcfg.getString("webBase") + helper.getTaskURL(task.getID()));

            String tplMailPath = Utils.exportDocument(mtpl, Conf.MainWFUpdate.MainPath, mtpn + "[" + uniqueId + "]");
            String mailExcelPath = Utils.saveDocReviewExcel(tplMailPath, Conf.MainWFUpdateSheetIndex.Mail,
                    Conf.MainWFUpdate.MainPath + "/" + mtpn + "[" + uniqueId + "].xlsx", dbks
            );
            String mailHtmlPath = Utils.convertExcelToHtml(mailExcelPath, Conf.MainWFUpdate.MainPath + "/" + mtpn + "[" + uniqueId + "].html");

            if(mlst.size() > 0) {
                JSONObject mail = new JSONObject();

                mail.put("To", String.join(";", mlst));
                mail.put("Subject", "WF-Task > " + dbks.getString("DocNo") + " / " + dbks.getString("RevNo"));
                mail.put("BodyHTMLFile", mailHtmlPath);

                try{
                    Utils.sendHTMLMail(ses, srv, mtpn, mail);
                } catch (Exception ex){
                    System.out.println("EXCP [Send-Mail] : " + ex.getMessage());
                }
            }

            System.out.println("Tested.");

        } catch (Exception e) {
            //throw new RuntimeException(e);
            System.out.println("Exception       : " + e.getMessage());
            System.out.println("    Class       : " + e.getClass());
            System.out.println("    Stack-Trace : " + e.getStackTrace() );
            return resultRestart("Exception : " + e.getMessage(),10);
        }

        System.out.println("Finished");
        return resultSuccess("Ended successfully");
    }
}