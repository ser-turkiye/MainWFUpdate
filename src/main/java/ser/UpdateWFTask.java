package ser;

import com.ser.blueline.IDocument;
import com.ser.blueline.IDocumentServer;
import com.ser.blueline.ISession;
import com.ser.blueline.bpm.*;
import de.ser.doxis4.agentserver.UnifiedAgent;

import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Collection;
import java.util.List;


public class UpdateWFTask extends UnifiedAgent {

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

            String taskCreation = (task.getCreationDate() == null ? "" : (new SimpleDateFormat("yyyyMMdd")).format(task.getCreationDate()));

            //task.getCurrentWorkbasket().getName();
            Collection<ITask> tasks = proi.findTasks(TaskStatus.READY);
            List<String> wlst = new ArrayList<>();
            for(ITask ftask : tasks){
                if(ftask.getCurrentWorkbasket() == null){continue;}
                wlst.add(ftask.getCurrentWorkbasket().getName());
            }

            IDocument mainDocument = (IDocument) proi.getMainInformationObject();

            mainDocument.setDescriptorValue("ccmPrjDocWFProcessName", "Main Document Review");
            mainDocument.setDescriptorValue("ccmPrjDocWFTaskName", task.getName());
            mainDocument.setDescriptorValue("ccmPrjDocWFTaskCreation", taskCreation);
            mainDocument.setDescriptorValue("ccmPrjDocWFTaskRecipients",
                String.join(";", wlst)
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