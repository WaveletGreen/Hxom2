package com.connor.HXom052;

import org.eclipse.core.commands.AbstractHandler;
import org.eclipse.core.commands.ExecutionEvent;
import org.eclipse.core.commands.ExecutionException;

import com.connor.Hxom.common.handlerUtil.CommonHandler;

public class ExportPPAPMatrixHandler extends AbstractHandler {
	@Override
	public Object execute(ExecutionEvent arg0) throws ExecutionException {
		CommonHandler action = new CommonHandler("com.connor.HXom052.ExportPPAPMatrixOperation");
		action.CallCommonAction();
		return null;
	}

}
