package ser;

import com.ser.blueline.IDocument;
import com.ser.blueline.IDocumentServer;
import com.ser.blueline.ISession;
import com.ser.blueline.bpm.IBpmService;
import com.ser.blueline.bpm.IProcessInstance;
import com.ser.blueline.bpm.IReceivers;
import com.ser.blueline.bpm.ITask;
import de.ser.doxis4.agentserver.UnifiedAgent;

import java.text.DateFormat;
import java.text.SimpleDateFormat;


public class UpdateWFProcIns extends UnifiedAgent {

    ISession ses;
    IDocumentServer srv;
    IBpmService bpm;
    private ProcessHelper helper;
    @Override
    protected Object execute() {
        if (getEventTask() == null)
            return resultError("Null Document object");

        try {
            ITask task = getEventTask();
            IProcessInstance proi = task.getProcessInstance();

            IDocument mainDocument = (IDocument) proi.getMainInformationObject();
            mainDocument.setDescriptorValue("ccmPrjDocWFProcessName", "Main Document Review");
            mainDocument.setDescriptorValue("ccmPrjDocWFTaskName", "Completed");
            mainDocument.setDescriptorValueTyped("ccmPrjDocWFTaskCreation", task.getCreationDate());
            mainDocument.setDescriptorValue("ccmPrjDocWFTaskRecipients",
               ""
            );

            mainDocument.commit();
            System.out.println("Tested.");

        } catch (Exception e) {
            //throw new RuntimeException(e);
            System.out.println("Exception       : " + e.getMessage());
            System.out.println("    Class       : " + e.getClass());
            System.out.println("    Stack-Trace : " + e.getStackTrace() );
            return resultError("Exception : " + e.getMessage());
        }

        System.out.println("Finished");
        return resultSuccess("Ended successfully");
    }
}