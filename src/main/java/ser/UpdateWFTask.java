package ser;

import com.ser.blueline.*;
import com.ser.blueline.bpm.*;
import de.ser.doxis4.agentserver.UnifiedAgent;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.json.JSONObject;

import java.io.File;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.concurrent.TimeUnit;


public class UpdateWFTask extends UnifiedAgent {
    Logger log = LogManager.getLogger();
    private ITask mainTask;
    private ProcessHelper helper;
    String approveCode = "";
    String decisionCode  = "";
    @Override
    protected Object execute() {
        if (getEventTask() == null)
            return resultError("Null Document object");

        log.info("Event Task starting..Task ID:" + getEventTask().getID());
        log.info("Event Task starting..Task Code:" + getEventTask().getCode());

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
            IProcessInstance proi = task.getProcessInstance();
            IDocument mainDocument = (IDocument) proi.getMainInformationObject();
            if(mainDocument == null){return resultSuccess("No-Main document");}

            log.info("Main Document is.." + mainDocument.getID());
            if(Objects.equals(task.getCode(), "Step04")){
                Collection<ITask> tsks = proi.findTasks();
                for(ITask ttsk : tsks) {
                    if (ttsk.getStatus() != TaskStatus.COMPLETED) {
                        continue;
                    }
                    if(decisionCode != ""){break;}
                    String tnam = (ttsk.getName() != null ? ttsk.getName() : "");
                    String tcod = (ttsk.getCode() != null ? ttsk.getCode() : "");
                    log.info("TASK-Name[" + tnam + "]");
                    log.info("TASK-Code[" + tcod + "]");
                    if(ttsk.getLoadedParentTask() != null && (tnam.equals("Consolidator Review") || tcod.equals("Step03"))){
                        decisionCode = ttsk.getDecision().getCode();
                        log.info("Approval Code updated.. decisionCode IS:" + decisionCode);
                        mainDocument.setDescriptorValue("ccmPrjDocApprCode", decisionCode);
                        log.info("Approval Code updated.. APPR CODE IS:" + mainDocument.getDescriptorValue("ccmPrjDocApprCode"));
                        //mainDocument.commit();
                        //log.info("UpdateWFTask maindoc committed..111");
                    }
                }
            }

            if(task.getAutoCompletionRule() != null){return resultSuccess("Auto Completion Rule ...");}

            IWorkbasket wbsk = task.getCurrentWorkbasket();
            if(wbsk == null){return resultSuccess("No-Workbasket ...");}

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
            if(Utils.hasDescriptor(proi, "ProcessID") && proi.getDescriptorValue("ProcessID") == null) {
                proi.setDescriptorValue("ProcessID", proi.getID());
                proi.commit();
            }
            if(Utils.hasDescriptor(task, "TaskID") && task.getDescriptorValue("TaskID") == null) {
                task.setDescriptorValue("TaskID", task.getID());
                task.commit();
            }
            mainDocument.commit();
            log.info("UpdateWFTask maindoc committed...222");


            Date tbgn = null, tend = new Date();
            if(task.getReadyDate() != null){
                tbgn = task.getReadyDate();
            }
            long durd  = 0L;
            double durh  = 0.0;
            if(tend != null && tbgn != null) {
                long diff = (tend.getTime() > tbgn.getTime() ? tend.getTime() - tbgn.getTime() : tbgn.getTime() - tend.getTime());
                durd = TimeUnit.DAYS.convert(diff, TimeUnit.MILLISECONDS);
                durh = ((TimeUnit.MINUTES.convert(diff, TimeUnit.MILLISECONDS) - (durd * 24 * 60)) * 100 / 60) / 100d;
            }
            String rcvf = "", rcvo = "";
            if(task.getPreviousWorkbasket() != null){
                rcvf = task.getPreviousWorkbasket().getFullName();
            }
            if(tbgn != null){
                rcvo = (new SimpleDateFormat("dd-MM-yyyy HH:mm")).format(tbgn);
            }
            String prjn = "",  mdno = "", mdrn = "", mdnm = "";
            if(mainDocument != null &&  Utils.hasDescriptor((IInformationObject) mainDocument, Conf.Descriptors.ProjectNo)){
                prjn = mainDocument.getDescriptorValue(Conf.Descriptors.ProjectNo, String.class);
            }
            if(mainDocument != null &&  Utils.hasDescriptor((IInformationObject) mainDocument, Conf.Descriptors.DocNumber)){
                mdno = mainDocument.getDescriptorValue(Conf.Descriptors.DocNumber, String.class);
            }
            if(mainDocument != null &&  Utils.hasDescriptor((IInformationObject) mainDocument, Conf.Descriptors.Revision)){
                mdrn = mainDocument.getDescriptorValue(Conf.Descriptors.Revision, String.class);
            }
            if(mainDocument != null &&  Utils.hasDescriptor((IInformationObject) mainDocument, Conf.Descriptors.Name)){
                mdnm = mainDocument.getDescriptorValue(Conf.Descriptors.Name, String.class);
            }



            this.helper = new ProcessHelper(Utils.session);

            (new File(Conf.Paths.MainPath)).mkdirs();

            //String prjn = proi.getDescriptorValue(Conf.Descriptors.ProjectNo, String.class);
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
            dbks.put("DocName", (mdnm != null  ? mdnm : ""));
            dbks.put("ReceivedOn", (rcvo != null ? rcvo : ""));
            dbks.put("ProcessTitle", (proi != null ? proi.getDisplayName() : ""));
            dbks.put("ProjectNo", (prjn != null  ? prjn : ""));


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
                //mail.put("Subject", "WF-Task > " + dbks.getString("DocNo") + " / " + dbks.getString("RevNo"));
                mail.put("Subject", "New Doc. > " + task.getName() + " - " + dbks.getString("DocNo") + " / " + dbks.getString("RevNo"));
                mail.put("BodyHTMLFile", mailHtmlPath);

                try{
                    Utils.sendHTMLMail(mail,null);
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