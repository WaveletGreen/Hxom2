package com.connor.timeTester;

import java.util.Map;

import com.teamcenter.rac.aif.AbstractAIFApplication;
import com.teamcenter.rac.aif.AbstractAIFOperation;
import com.teamcenter.rac.aifrcp.AIFUtility;
import com.teamcenter.rac.kernel.TCComponent;
import com.teamcenter.rac.kernel.TCComponentBOMWindow;
import com.teamcenter.rac.kernel.TCComponentScheduleTask;
import com.teamcenter.rac.kernel.TCSession;

public class ScheduleTaskOperation extends AbstractAIFOperation {

	private TCSession session;
	private AbstractAIFApplication app;

	public ScheduleTaskOperation() {
		super();
		this.app = AIFUtility.getCurrentApplication();
		this.session = (TCSession) app.getSession();
	}

	@Override
	public void executeOperation() throws Exception {
		TCComponent component = (TCComponent) app.getTargetComponent();
		String type = component.getProperty("object_type");
		TCComponentScheduleTask task = null;
		if (component instanceof TCComponentScheduleTask) {
			task = (TCComponentScheduleTask) component;
		}
		Map<String, String> prop = task.getProperties();
		for (String key : prop.keySet()) {
			System.out.println(key + "----" + prop.get(key));
		}
		System.out.println();
	}

}
