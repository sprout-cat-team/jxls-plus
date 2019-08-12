package com.tang.jxlsplus.formula;

import org.jxls.area.Area;
import org.jxls.formula.AbstractFormulaProcessor;
import org.jxls.formula.FastFormulaProcessor;
import org.jxls.transform.Transformer;

/**
 * 自定义excel公式处理器 （解决 原始格式处理器迭代行数过多时出错的问题）；<br>
 * 参照 {@link FastFormulaProcessor}
 * @author tzg
 */
public class JxlsPlusFormulaProcessor extends AbstractFormulaProcessor {

    @Override
    public void processAreaFormulas(Transformer transformer, Area area) {

    }
}
