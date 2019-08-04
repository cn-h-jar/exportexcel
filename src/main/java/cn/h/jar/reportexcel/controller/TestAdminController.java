package cn.h.jar.reportexcel.controller;

import cn.h.jar.reportexcel.utils.ExcelUtils;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletResponse;
import java.util.*;

@RestController
@RequestMapping("api/test")
public class TestAdminController {

    /**
     * 导出分页列表的excle 接口
     *
     * @return
     */
    @RequestMapping("exportExcel")
    public void exportExcel(@RequestParam HashMap map, HttpServletResponse response) throws Exception {

        List<Map<String,Object>> mapList=new ArrayList<>();
        Map<String,Object> mapRow;

        mapRow=new LinkedHashMap<>();
        mapRow.put("name","姓名");
        mapRow.put("age","年龄");
        mapRow.put("marry","婚否");
        mapList.add(mapRow);

        mapRow=new LinkedHashMap<>();
        mapRow.put("name","CDN");
        mapRow.put("age","60");
        mapRow.put("marry","重婚");
        mapList.add(mapRow);

        mapRow=new LinkedHashMap<>();
        mapRow.put("name","John");
        mapRow.put("age","70");
        mapRow.put("marry","多次离异");
        mapList.add(mapRow);

        mapRow=new LinkedHashMap<>();
        mapRow.put("name","Jar");
        mapRow.put("age","18");
        mapRow.put("marry","未婚");
        mapList.add(mapRow);

        ExcelUtils.generateWrite(mapList,"婚姻状况","组1",response);
    }

}
