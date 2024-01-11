package ser;

import com.ser.blueline.IDocument;
import com.ser.blueline.IDocumentServer;
import com.ser.blueline.IMutability;
import com.ser.blueline.ISession;
import com.ser.blueline.bpm.IBpmService;
import com.ser.blueline.bpm.IProcessInstance;
import com.ser.blueline.bpm.IReceivers;
import com.ser.blueline.bpm.ITask;
import de.ser.doxis4.agentserver.UnifiedAgent;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Objects;


public class UpdateWFProcIns extends UnifiedAgent {
    Logger log = LogManager.getLogger();
    private ProcessHelper processHelper;
    @Override
    protected Object execute() {
        if (getEventTask() == null)
            return resultError("Null Document object");

        Utils.session = getSes();
        Utils.bpm = getBpm();
        Utils.server = Utils.session.getDocumentServer();
        Utils.loadDirectory(Conf.Paths.MainPath);

        try {
            processHelper = new ProcessHelper(Utils.session);

            ITask task = getEventTask();
            IProcessInstance proi = task.getProcessInstance();

            IDocument mainDocument = (IDocument) proi.getMainInformationObject();
            String chck = proi.getDescriptorValue("ccmCrrsStatus", String.class);

            String stts = "Completed";
            if(chck != null && chck.equals("Cancelled")){
                stts = chck;
                mainDocument.setDescriptorValue("ccmPrjDocApprCode", "-");
            }


            mainDocument.setDescriptorValue("ccmPrjDocWFProcessName", "Main Document Review");
            mainDocument.setDescriptorValue("ccmPrjDocWFTaskName", stts);
            mainDocument.setDescriptorValueTyped("ccmPrjDocWFTaskCreation", task.getCreationDate());
            mainDocument.setDescriptorValue("ccmPrjDocWFTaskRecipients",
               ""
            );

            mainDocument.commit();

            //gecisi kaldirildi
           /* if(mainDocument.getMutability() != IMutability.IMMUTABLE) {
                mainDocument.commit();
                getSes().getDocumentServer().updateMutability(getSes(), mainDocument, IMutability.IMMUTABLE);
            }*/

            log.info("Tested.");

        } catch (Exception e) {
            //throw new RuntimeException(e);
            log.error("Exception       : " + e.getMessage());
            log.error("    Class       : " + e.getClass());
            log.error("    Stack-Trace : " + e.getStackTrace() );
            return resultRestart("Exception : " + e.getMessage(), 10);

        }

        log.info("Finished");
        return resultSuccess("Ended successfully");
    }
}