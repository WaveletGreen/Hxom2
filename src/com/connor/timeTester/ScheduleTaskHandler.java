package com.connor.timeTester;

import org.eclipse.core.commands.AbstractHandler;
import org.eclipse.core.commands.ExecutionEvent;
import org.eclipse.core.commands.ExecutionException;

import com.connor.Hxom.common.handlerUtil.CommonHandler;

public class ScheduleTaskHandler extends AbstractHandler {
	@Override
	public Object execute(ExecutionEvent arg0) throws ExecutionException {
		CommonHandler action = new CommonHandler("com.connor.timeTester.ScheduleTaskOperation");
		action.CallCommonAction();
		return null;
	}

}
