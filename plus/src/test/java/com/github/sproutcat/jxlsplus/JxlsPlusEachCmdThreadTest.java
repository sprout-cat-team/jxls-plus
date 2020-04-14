package com.github.sproutcat.jxlsplus;

import lombok.extern.slf4j.Slf4j;

/**
 * each 指令多线程测试用例
 */
@Slf4j
public class JxlsPlusEachCmdThreadTest {

    public static void main(String[] args) {
        JxlsPlusTestData testData = JxlsPlusTestData.init();

        log.debug("eachCmd 多线程测试");
        String template = "/each_template.xlsx";
        for (int i = 0; i < 100; i++) {
            String fileName = "jp_eachDemo_thread_" + (i + 1);
            Thread myThread = new Thread(() -> {

                log.debug("==>> fileName: {}", fileName);
                JxlsPlusUtils.processTemplate(
                        JxlsPlusTestData.class.getResourceAsStream(template),
                        String.format("C:\\Users\\%s\\Desktop\\jxls-plus\\%s.xlsx", testData.getCurrentUser(), fileName),
                        testData.getContext()
                );
            }, fileName);

            myThread.start();
        }
    }

}
