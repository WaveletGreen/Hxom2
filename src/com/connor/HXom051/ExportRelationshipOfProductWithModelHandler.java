package com.connor.HXom051;

import org.eclipse.core.commands.AbstractHandler;
import org.eclipse.core.commands.ExecutionEvent;
import org.eclipse.core.commands.ExecutionException;

import com.connor.Hxom.common.handlerUtil.CommonHandler;

public class ExportRelationshipOfProductWithModelHandler extends AbstractHandler {
	@Override
	public Object execute(ExecutionEvent arg0) throws ExecutionException {
		CommonHandler action = new CommonHandler("com.connor.HXom051.ExportRelationshipOfProductWithModelOperation");
		action.CallCommonAction();
		return null;
	}

}
