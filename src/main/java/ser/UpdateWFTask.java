package ser;

import com.ser.blueline.*;
import com.ser.blueline.bpm.*;
import de.ser.doxis4.agentserver.UnifiedAgent;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.json.JSONObject;

import java.io.File;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;
import java.util.UUID;


public class UpdateWFTask extends UnifiedAgent {
    Logger log = LogManager.getLogger();
    private ProcessHelper helper;
    @Override
    protected Object execute() {
        if (getEventTask() == null)
            return resultError("Null Document object");


        Utils.session = getSes();
        Utils.bpm = getBpm();
        Utils.server = Utils.session.getDocumentServer();
        Utils.loadDirectory(Conf.Paths.MainPath);

        try {

            JSONObject scfg = Utils.getSystemConfig();
            if(scfg.has("LICS.SPIRE_XLS")){
                com.spire.license.LicenseProvider.setLicenseKey(scfg.getString("LICS.SPIRE_XLS"));
            }

            ITask task = getEventTask();
            if(task.getAutoCompletionRule() != null){return resultSuccess("Auto Completion Rule ...");}

            IWorkbasket wbsk = task.getCurrentWorkbasket();
            if(wbsk == null){return resultSuccess("No-Workbasket ...");}

            IProcessInstance proi = task.getProcessInstance();

            List<String> wlst = new ArrayList<>();
            List<String> mlst = new ArrayList<>();
            if(wbsk != null) {
                if (!wbsk.getName().isEmpty()) {
                    wlst.add(wbsk.getName());
                }
                String wuem = wbsk.getNotifyEMail();
                if (wuem != null && !wuem.isEmpty()) {
                    mlst.add(wuem);
                }
            }

            IDocument mainDocument = (IDocument) proi.getMainInformationObject();
            if(mainDocument == null){return resultSuccess("No-Main document");}

            String docNo = mainDocument.getDescriptorValue(Conf.Descriptors.DocNumber, String.class);
            docNo = (docNo == null ? "" : docNo);
            if(docNo.isEmpty()){return resultSuccess("Passed successfully");}

            String revNo = mainDocument.getDescriptorValue(Conf.Descriptors.Revision, String.class);
            revNo = (revNo == null ? "" : revNo);

            String taskCreation = (task.getCreationDate() == null ? "" : (new SimpleDateFormat("yyyyMMdd")).format(task.getCreationDate()));

            mainDocument.setDescriptorValue("ccmPrjDocWFProcessName", proi.getDisplayName());
            mainDocument.setDescriptorValue("ccmPrjDocWFTaskName", task.getName());
            mainDocument.setDescriptorValue("ccmPrjDocWFTaskCreation", taskCreation);
            mainDocument.setDescriptorValue("ccmPrjDocWFTaskRecipients",
                    String.join(";", wlst)
            );
            mainDocument.commit();

            this.helper = new ProcessHelper(Utils.session);

            (new File(Conf.Paths.MainPath)).mkdirs();

            String prjn = proi.getDescriptorValue(Conf.Descriptors.ProjectNo, String.class);
            if(prjn.isEmpty()){
                throw new Exception("Project no is empty.");
            }
            IInformationObject prjt = Utils.getProjectWorkspace(prjn, helper);
            if(prjt == null){
                throw new Exception("Project not found [" + prjn + "].");
            }

            if(mlst.size() == 0){return resultSuccess("No mail address : " + (wbsk !=null ? wbsk.getFullName() : "-No Workbasket-"));}

            String uniqueId = UUID.randomUUID().toString();

            String mtpn = "UPDATE_WF_MAIL";
            IDocument mtpl = Utils.getTemplateDocument(prjt, mtpn);
            if(mtpl == null){
                return resultSuccess("No-Mail Template");
            }

            JSONObject dbks = new JSONObject();
            dbks.put("DocNo", docNo);
            dbks.put("RevNo", revNo);
            dbks.put("Title", mainDocument.getDisplayName());
            dbks.put("Task", task.getName());


            JSONObject mcfg = Utils.getMailConfig();
            dbks.put("DoxisLink", mcfg.getString("webBase") + helper.getTaskURL(task.getID()));

            String tplMailPath = Utils.exportDocument(mtpl, Conf.Paths.MainPath, mtpn + "[" + uniqueId + "]");
            String mailExcelPath = Utils.saveDocReviewExcel(tplMailPath, Conf.MainWFUpdateSheetIndex.Mail,
                    Conf.Paths.MainPath + "/" + mtpn + "[" + uniqueId + "].xlsx", dbks
            );
            String mailHtmlPath = Utils.convertExcelToHtml(mailExcelPath, Conf.Paths.MainPath + "/" + mtpn + "[" + uniqueId + "].html");

            if(mlst.size() > 0) {
                JSONObject mail = new JSONObject();

                mail.put("To", String.join(";", mlst));
                mail.put("Subject", "WF-Task > " + dbks.getString("DocNo") + " / " + dbks.getString("RevNo"));
                mail.put("BodyHTMLFile", mailHtmlPath);

                try{
                    Utils.sendHTMLMail(mail);
                } catch (Exception ex){
                    log.info("EXCP [Send-Mail] : " + ex.getMessage());
                }
            }

            log.info("Tested.");

        } catch (Exception e) {
            //throw new RuntimeException(e);
            log.error("Exception       : " + e.getMessage());
            log.error("    Class       : " + e.getClass());
            log.error("    Stack-Trace : " + e.getStackTrace() );
            return resultRestart("Exception : " + e.getMessage(),10);
        }

        log.info("Finished");
        return resultSuccess("Ended successfully");
    }
}