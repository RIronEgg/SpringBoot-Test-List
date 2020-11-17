package com.delsql.ironegg.Excel;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.TypeReference;
import com.delsql.ironegg.Excel.bean.NewContractStatisticsTableDco;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.springframework.beans.BeanUtils;

import java.io.*;
import java.util.*;
import java.util.regex.Pattern;

/**
 * @author RenZiHan
 * @version 1.0
 * @date 2020/11/17 4:02 下午
 * @purpose
 */
public class ExcelTest {
    public static void main(String[] args) {
        //EXcel数据
        String json = "[{\"deduction\":[\"YLZHT\",\"YLFBHT\",\"NSZHT\"],\"level\":1,\"orgCode\":[\"CHIN_160000\",\"CHIN_360000\",\"BJGC_0101020000\",\"YNGW_0102070000\",\"BJGW_0102010000\",\"ZJGW_0102020000\",\"SXGW_0102030000\",\"HNGW_0102040000\",\"SCGW_0102050000\",\"GZGW_0102060000\",\"SCKF_0101210000\",\"BJNY_0101240000\",\"GWJT_0102140000\",\"GSJS_0101270000\",\"BJHT_0101420000\",\"BJLQ_0101200000\",\"BJTL_0101230000\",\"SHJZ_0101330000\",\"BJGJ_0101190000\",\"DJJJ_0109000000\",\"CWGS_0101320000\",\"DJBL_0111000000\",\"BJHK_0101280000\",\"HBYJ_0105010000\",\"SDHD_0105030000\",\"SDYG_0105040000\",\"JLSJ_0104020000\",\"SHSJ_0104030000\",\"FJSJ_0104040000\",\"HNKC_0104060000\",\"JXSJ_0104110000\",\"SCKC_0104070000\",\"QHSJ_0104120000\",\"HBSJ_0104010000\",\"CQJS_0105200000\",\"SDEG_0105050000\",\"SDSG_0105060000\",\"SHJS_0105080000\",\"HBGC_0101500000\",\"HNEJ_0105150000\",\"JXHD_0105160000\",\"JXSD_0105170000\",\"GZYG_0105230000\",\"GZSJ_0104090000\",\"BJFC_0101220000\",\"JLGC_0101010000\",\"SXGC_0101030000\",\"CHGC_0101040000\",\"SCGC_0101050000\",\"LNGC_0101060000\",\"SCGC_0101070000\",\"HNGC_0101080000\",\"GZGC_0101090000\",\"SCGC_0101100000\",\"HNGC_0101110000\",\"ZJGC_0101120000\",\"TJGC_0101130000\",\"YNGC_0101140000\",\"SXGC_0101150000\",\"DJDG_0106210000\",\"CCSB_0106030000\",\"CDJJ_0106170000\",\"HBSB_0106010000\",\"JXSB_0106130000\",\"SHXZ_0106060000\",\"SCJX_0106150000\",\"HKSB_0106090000\",\"WHSB_0106080000\",\"ZBYJ_0106200000\",\"FJGC_0101160000\",\"TJJC_0101170000\",\"TJHG_0101260000\",\"BJZP_0101250000\"],\"orgName\":\"电建集团（抵扣后）\",\"reportTime\":\"2020-11-12\"},{\"deduction\":[\"YLZHT\",\"YLFBHT\",\"NSZHT\",\"NSFBHT\"],\"level\":1,\"orgCode\":[\"CHIN_160000\",\"CHIN_360000\",\"BJGC_0101020000\",\"YNGW_0102070000\",\"BJGW_0102010000\",\"ZJGW_0102020000\",\"SXGW_0102030000\",\"HNGW_0102040000\",\"SCGW_0102050000\",\"GZGW_0102060000\",\"SCKF_0101210000\",\"BJNY_0101240000\",\"GWJT_0102140000\",\"GSJS_0101270000\",\"BJHT_0101420000\",\"BJLQ_0101200000\",\"BJTL_0101230000\",\"SHJZ_0101330000\",\"BJGJ_0101190000\",\"DJJJ_0109000000\",\"CWGS_0101320000\",\"DJBL_0111000000\",\"BJHK_0101280000\",\"HBYJ_0105010000\",\"SDHD_0105030000\",\"SDYG_0105040000\",\"JLSJ_0104020000\",\"SHSJ_0104030000\",\"FJSJ_0104040000\",\"HNKC_0104060000\",\"JXSJ_0104110000\",\"SCKC_0104070000\",\"QHSJ_0104120000\",\"HBSJ_0104010000\",\"CQJS_0105200000\",\"SDEG_0105050000\",\"SDSG_0105060000\",\"SHJS_0105080000\",\"HBGC_0101500000\",\"HNEJ_0105150000\",\"JXHD_0105160000\",\"JXSD_0105170000\",\"GZYG_0105230000\",\"GZSJ_0104090000\",\"BJFC_0101220000\",\"JLGC_0101010000\",\"SXGC_0101030000\",\"CHGC_0101040000\",\"SCGC_0101050000\",\"LNGC_0101060000\",\"SCGC_0101070000\",\"HNGC_0101080000\",\"GZGC_0101090000\",\"SCGC_0101100000\",\"HNGC_0101110000\",\"ZJGC_0101120000\",\"TJGC_0101130000\",\"YNGC_0101140000\",\"SXGC_0101150000\",\"DJDG_0106210000\",\"CCSB_0106030000\",\"CDJJ_0106170000\",\"HBSB_0106010000\",\"JXSB_0106130000\",\"SHXZ_0106060000\",\"SCJX_0106150000\",\"HKSB_0106090000\",\"WHSB_0106080000\",\"ZBYJ_0106200000\",\"FJGC_0101160000\",\"TJJC_0101170000\",\"TJHG_0101260000\",\"BJZP_0101250000\"],\"orgName\":\"电建集团（抵扣前）\",\"reportTime\":\"2020-11-12\"},{\"deduction\":[\"YLZHT\",\"YLFBHT\",\"NSZHT\",\"NSFBHT\"],\"level\":1,\"orgCode\":[\"CHIN_280000\"],\"orgName\":\"规划总院\",\"reportTime\":\"2020-11-12\"},{\"deduction\":[\"YLZHT\",\"YLFBHT\",\"NSZHT\",\"NSFBHT\"],\"level\":1,\"orgCode\":[\"SCKF_0101210000\",\"BJNY_0101240000\",\"GWJT_0102140000\",\"GSJS_0101270000\",\"BJHT_0101420000\",\"BJLQ_0101200000\",\"BJTL_0101230000\",\"SHJZ_0101330000\",\"BJGJ_0101190000\",\"DJJJ_0109000000\",\"CWGS_0101320000\",\"DJBL_0111000000\",\"BJFC_0101220000\",\"BJZP_0101250000\"],\"orgName\":\"平台公司\",\"reportTime\":\"2020-11-12\"},{\"deduction\":[\"YLZHT\",\"YLFBHT\",\"NSZHT\",\"NSFBHT\"],\"level\":1,\"orgCode\":[\"SCKF_0101210000\",\"BJNY_0101240000\",\"GWJT_0102140000\",\"GSJS_0101270000\",\"BJHT_0101420000\",\"BJFC_0101220000\"],\"orgName\":\"平台公司（投资运营类）\",\"reportTime\":\"2020-11-12\"},{\"deduction\":[\"YLZHT\",\"YLFBHT\",\"NSZHT\",\"NSFBHT\"],\"level\":2,\"orgCode\":[\"SCKF_0101210000\"],\"orgName\":\"电建水电开发公司\",\"reportTime\":\"2020-11-12\"},{\"deduction\":[\"YLZHT\",\"YLFBHT\",\"NSZHT\",\"NSFBHT\"],\"level\":2,\"orgCode\":[\"BJNY_0101240000\"],\"orgName\":\"水电新能源公司\",\"reportTime\":\"2020-11-12\"},{\"deduction\":[\"YLZHT\",\"YLFBHT\",\"NSZHT\",\"NSFBHT\"],\"level\":2,\"orgCode\":[\"GWJT_0102140000\"],\"orgName\":\"水电顾问公司\",\"reportTime\":\"2020-11-12\"},{\"deduction\":[\"YLZHT\",\"YLFBHT\",\"NSZHT\",\"NSFBHT\"],\"level\":2,\"orgCode\":[\"GSJS_0101270000\"],\"orgName\":\"甘肃能源公司\",\"reportTime\":\"2020-11-12\"},{\"deduction\":[\"YLZHT\",\"YLFBHT\",\"NSZHT\",\"NSFBHT\"],\"level\":2,\"orgCode\":[\"BJHT_0101420000\"],\"orgName\":\"电建海投公司\",\"reportTime\":\"2020-11-12\"},{\"deduction\":[\"YLZHT\",\"YLFBHT\",\"NSZHT\",\"NSFBHT\"],\"level\":2,\"orgCode\":[\"BJFC_0101220000\"],\"orgName\":\"电建地产公司\",\"reportTime\":\"2020-11-12\"},{\"deduction\":[\"YLZHT\",\"YLFBHT\",\"NSZHT\",\"NSFBHT\"],\"level\":1,\"orgCode\":[\"BJHK_0101280000\"],\"orgName\":\"其它公司\",\"reportTime\":\"2020-11-12\"},{\"deduction\":[\"YLZHT\",\"YLFBHT\",\"NSZHT\",\"NSFBHT\"],\"level\":2,\"orgCode\":[\"BJHK_0101280000\"],\"orgName\":\"华科软公司\",\"reportTime\":\"2020-11-12\"},{\"deduction\":[\"YLZHT\",\"YLFBHT\",\"NSZHT\",\"NSFBHT\"],\"level\":1,\"orgCode\":[\"BJLQ_0101200000\",\"BJTL_0101230000\",\"SHJZ_0101330000\",\"BJGJ_0101190000\"],\"orgName\":\"平台公司（专业市场类）\",\"reportTime\":\"2020-11-12\"},{\"deduction\":[\"YLZHT\",\"YLFBHT\",\"NSZHT\",\"NSFBHT\"],\"level\":2,\"orgCode\":[\"BJLQ_0101200000\"],\"orgName\":\"电建路桥公司\",\"reportTime\":\"2020-11-12\"},{\"deduction\":[\"YLZHT\",\"YLFBHT\",\"NSZHT\",\"NSFBHT\"],\"level\":2,\"orgCode\":[\"BJTL_0101230000\"],\"orgName\":\"电建铁路公司\",\"reportTime\":\"2020-11-12\"},{\"deduction\":[\"YLZHT\",\"YLFBHT\",\"NSZHT\",\"NSFBHT\"],\"level\":2,\"orgCode\":[\"SHJZ_0101330000\"],\"orgName\":\"电建生态公司\",\"reportTime\":\"2020-11-12\"},{\"deduction\":[\"YLZHT\",\"YLFBHT\",\"NSZHT\",\"NSFBHT\"],\"level\":2,\"orgCode\":[\"BJGJ_0101190000\"],\"orgName\":\"电建国际公司 \",\"reportTime\":\"2020-11-12\"},{\"deduction\":[\"YLZHT\",\"YLFBHT\",\"NSZHT\",\"NSFBHT\"],\"level\":1,\"orgCode\":[\"JLSJ_0104020000\",\"SHSJ_0104030000\",\"FJSJ_0104040000\",\"HNKC_0104060000\",\"JXSJ_0104110000\",\"SCKC_0104070000\",\"QHSJ_0104120000\",\"HBSJ_0104010000\",\"GZSJ_0104090000\"],\"orgName\":\"电力设计\",\"reportTime\":\"2020-11-12\"},{\"deduction\":[\"YLZHT\",\"YLFBHT\",\"NSZHT\",\"NSFBHT\"],\"level\":2,\"orgCode\":[\"JLSJ_0104020000\"],\"orgName\":\"吉林院\",\"reportTime\":\"2020-11-12\"},{\"deduction\":[\"YLZHT\",\"YLFBHT\",\"NSZHT\",\"NSFBHT\"],\"level\":2,\"orgCode\":[\"SHSJ_0104030000\"],\"orgName\":\"上海院\",\"reportTime\":\"2020-11-12\"},{\"deduction\":[\"YLZHT\",\"YLFBHT\",\"NSZHT\",\"NSFBHT\"],\"level\":2,\"orgCode\":[\"FJSJ_0104040000\"],\"orgName\":\"福建院\",\"reportTime\":\"2020-11-12\"},{\"deduction\":[\"YLZHT\",\"YLFBHT\",\"NSZHT\",\"NSFBHT\"],\"level\":2,\"orgCode\":[\"HNKC_0104060000\"],\"orgName\":\"华中院\",\"reportTime\":\"2020-11-12\"},{\"deduction\":[\"YLZHT\",\"YLFBHT\",\"NSZHT\",\"NSFBHT\"],\"level\":2,\"orgCode\":[\"JXSJ_0104110000\"],\"orgName\":\"江西院\",\"reportTime\":\"2020-11-12\"},{\"deduction\":[\"YLZHT\",\"YLFBHT\",\"NSZHT\",\"NSFBHT\"],\"level\":2,\"orgCode\":[\"SCKC_0104070000\"],\"orgName\":\"四川设计咨询公司\",\"reportTime\":\"2020-11-12\"},{\"deduction\":[\"YLZHT\",\"YLFBHT\",\"NSZHT\",\"NSFBHT\"],\"level\":2,\"orgCode\":[\"QHSJ_0104120000\"],\"orgName\":\"青海院\",\"reportTime\":\"2020-11-12\"},{\"deduction\":[\"YLZHT\",\"YLFBHT\",\"NSZHT\",\"NSFBHT\"],\"level\":2,\"orgCode\":[\"HBSJ_0104010000\"],\"orgName\":\"河北院\",\"reportTime\":\"2020-11-12\"},{\"deduction\":[\"YLZHT\",\"YLFBHT\",\"NSZHT\",\"NSFBHT\"],\"level\":2,\"orgCode\":[\"GZSJ_0104090000\"],\"orgName\":\"贵州院\",\"reportTime\":\"2020-11-12\"},{\"deduction\":[\"YLZHT\",\"YLFBHT\",\"NSZHT\",\"NSFBHT\"],\"level\":1,\"orgCode\":[\"HBYJ_0105010000\",\"SDHD_0105030000\",\"SDYG_0105040000\",\"CQJS_0105200000\",\"SDEG_0105050000\",\"SDSG_0105060000\",\"SHJS_0105080000\",\"HBGC_0101500000\",\"HNEJ_0105150000\",\"JXHD_0105160000\",\"JXSD_0105170000\",\"GZYG_0105230000\"],\"orgName\":\"电力工程\",\"reportTime\":\"2020-11-12\"},{\"deduction\":[\"YLZHT\",\"YLFBHT\",\"NSZHT\",\"NSFBHT\"],\"level\":2,\"orgCode\":[\"HBYJ_0105010000\"],\"orgName\":\"河北工程公司\",\"reportTime\":\"2020-11-12\"},{\"deduction\":[\"YLZHT\",\"YLFBHT\",\"NSZHT\",\"NSFBHT\"],\"level\":2,\"orgCode\":[\"SDHD_0105030000\"],\"orgName\":\"山东电建公司\",\"reportTime\":\"2020-11-12\"},{\"deduction\":[\"YLZHT\",\"YLFBHT\",\"NSZHT\",\"NSFBHT\"],\"level\":2,\"orgCode\":[\"SDYG_0105040000\"],\"orgName\":\"山东电建一公司\",\"reportTime\":\"2020-11-12\"},{\"deduction\":[\"YLZHT\",\"YLFBHT\",\"NSZHT\",\"NSFBHT\"],\"level\":2,\"orgCode\":[\"CQJS_0105200000\"],\"orgName\":\"重庆工程公司\",\"reportTime\":\"2020-11-12\"},{\"deduction\":[\"YLZHT\",\"YLFBHT\",\"NSZHT\",\"NSFBHT\"],\"level\":2,\"orgCode\":[\"SDEG_0105050000\"],\"orgName\":\"电建核电公司\",\"reportTime\":\"2020-11-12\"},{\"deduction\":[\"YLZHT\",\"YLFBHT\",\"NSZHT\",\"NSFBHT\"],\"level\":2,\"orgCode\":[\"SDSG_0105060000\"],\"orgName\":\"山东电建三公司\",\"reportTime\":\"2020-11-12\"},{\"deduction\":[\"YLZHT\",\"YLFBHT\",\"NSZHT\",\"NSFBHT\"],\"level\":2,\"orgCode\":[\"SHJS_0105080000\"],\"orgName\":\"上海电建公司\",\"reportTime\":\"2020-11-12\"},{\"deduction\":[\"YLZHT\",\"YLFBHT\",\"NSZHT\",\"NSFBHT\"],\"level\":2,\"orgCode\":[\"HBGC_0101500000\"],\"orgName\":\"湖北工程公司\",\"reportTime\":\"2020-11-12\"},{\"deduction\":[\"YLZHT\",\"YLFBHT\",\"NSZHT\",\"NSFBHT\"],\"level\":2,\"orgCode\":[\"HNEJ_0105150000\"],\"orgName\":\"河南工程公司\",\"reportTime\":\"2020-11-12\"},{\"deduction\":[\"YLZHT\",\"YLFBHT\",\"NSZHT\",\"NSFBHT\"],\"level\":2,\"orgCode\":[\"JXHD_0105160000\"],\"orgName\":\"江西电建公司\",\"reportTime\":\"2020-11-12\"},{\"deduction\":[\"YLZHT\",\"YLFBHT\",\"NSZHT\",\"NSFBHT\"],\"level\":2,\"orgCode\":[\"JXSD_0105170000\"],\"orgName\":\"江西水电公司\",\"reportTime\":\"2020-11-12\"},{\"deduction\":[\"YLZHT\",\"YLFBHT\",\"NSZHT\",\"NSFBHT\"],\"level\":2,\"orgCode\":[\"GZYG_0105230000\"],\"orgName\":\"贵州工程公司\",\"reportTime\":\"2020-11-12\"},{\"deduction\":[\"YLZHT\",\"YLFBHT\",\"NSZHT\",\"NSFBHT\"],\"level\":1,\"orgCode\":[\"DJJJ_0109000000\",\"CWGS_0101320000\",\"DJBL_0111000000\",\"BJZP_0101250000\"],\"orgName\":\"平台公司（金融服务类）\",\"reportTime\":\"2020-11-12\"},{\"deduction\":[\"YLZHT\",\"YLFBHT\",\"NSZHT\",\"NSFBHT\"],\"level\":2,\"orgCode\":[\"DJJJ_0109000000\"],\"orgName\":\"电建基金公司\",\"reportTime\":\"2020-11-12\"},{\"deduction\":[\"YLZHT\",\"YLFBHT\",\"NSZHT\",\"NSFBHT\"],\"level\":2,\"orgCode\":[\"CWGS_0101320000\"],\"orgName\":\"电建财务公司\",\"reportTime\":\"2020-11-12\"},{\"deduction\":[\"YLZHT\",\"YLFBHT\",\"NSZHT\",\"NSFBHT\"],\"level\":2,\"orgCode\":[\"DJBL_0111000000\"],\"orgName\":\"电建保理公司\",\"reportTime\":\"2020-11-12\"},{\"deduction\":[\"YLZHT\",\"YLFBHT\",\"NSZHT\",\"NSFBHT\"],\"level\":2,\"orgCode\":[\"BJZP_0101250000\"],\"orgName\":\"电建租赁公司\",\"reportTime\":\"2020-11-12\"},{\"deduction\":[\"YLZHT\",\"YLFBHT\",\"NSZHT\",\"NSFBHT\"],\"level\":1,\"orgCode\":[\"YNGW_0102070000\",\"BJGW_0102010000\",\"ZJGW_0102020000\",\"SXGW_0102030000\",\"HNGW_0102040000\",\"SCGW_0102050000\",\"GZGW_0102060000\"],\"orgName\":\"水电设计\",\"reportTime\":\"2020-11-12\"},{\"deduction\":[\"YLZHT\",\"YLFBHT\",\"NSZHT\",\"NSFBHT\"],\"level\":2,\"orgCode\":[\"YNGW_0102070000\"],\"orgName\":\"昆明院\",\"reportTime\":\"2020-11-12\"},{\"deduction\":[\"YLZHT\",\"YLFBHT\",\"NSZHT\",\"NSFBHT\"],\"level\":2,\"orgCode\":[\"BJGW_0102010000\"],\"orgName\":\"北京院\",\"reportTime\":\"2020-11-12\"},{\"deduction\":[\"YLZHT\",\"YLFBHT\",\"NSZHT\",\"NSFBHT\"],\"level\":2,\"orgCode\":[\"ZJGW_0102020000\"],\"orgName\":\"华东院\",\"reportTime\":\"2020-11-12\"},{\"deduction\":[\"YLZHT\",\"YLFBHT\",\"NSZHT\",\"NSFBHT\"],\"level\":2,\"orgCode\":[\"SXGW_0102030000\"],\"orgName\":\"西北院\",\"reportTime\":\"2020-11-12\"},{\"deduction\":[\"YLZHT\",\"YLFBHT\",\"NSZHT\",\"NSFBHT\"],\"level\":2,\"orgCode\":[\"HNGW_0102040000\"],\"orgName\":\"中南院\",\"reportTime\":\"2020-11-12\"},{\"deduction\":[\"YLZHT\",\"YLFBHT\",\"NSZHT\",\"NSFBHT\"],\"level\":2,\"orgCode\":[\"SCGW_0102050000\"],\"orgName\":\"成都院\",\"reportTime\":\"2020-11-12\"},{\"deduction\":[\"YLZHT\",\"YLFBHT\",\"NSZHT\",\"NSFBHT\"],\"level\":2,\"orgCode\":[\"GZGW_0102060000\"],\"orgName\":\"贵阳院\",\"reportTime\":\"2020-11-12\"},{\"deduction\":[\"YLZHT\",\"YLFBHT\",\"NSZHT\",\"NSFBHT\"],\"level\":1,\"orgCode\":[\"DJDG_0106210000\",\"CCSB_0106030000\",\"CDJJ_0106170000\",\"HBSB_0106010000\",\"JXSB_0106130000\",\"SHXZ_0106060000\",\"SCJX_0106150000\",\"HKSB_0106090000\",\"WHSB_0106080000\",\"ZBYJ_0106200000\"],\"orgName\":\"装备制造\",\"reportTime\":\"2020-11-12\"},{\"deduction\":[\"YLZHT\",\"YLFBHT\",\"NSZHT\",\"NSFBHT\"],\"level\":2,\"orgCode\":[\"DJDG_0106210000\"],\"orgName\":\"电建德国公司\",\"reportTime\":\"2020-11-12\"},{\"deduction\":[\"YLZHT\",\"YLFBHT\",\"NSZHT\",\"NSFBHT\"],\"level\":2,\"orgCode\":[\"CCSB_0106030000\"],\"orgName\":\"长春设备公司\",\"reportTime\":\"2020-11-12\"},{\"deduction\":[\"YLZHT\",\"YLFBHT\",\"NSZHT\",\"NSFBHT\"],\"level\":2,\"orgCode\":[\"CDJJ_0106170000\"],\"orgName\":\"电建器材公司\",\"reportTime\":\"2020-11-12\"},{\"deduction\":[\"YLZHT\",\"YLFBHT\",\"NSZHT\",\"NSFBHT\"],\"level\":2,\"orgCode\":[\"HBSB_0106010000\"],\"orgName\":\"河北装备公司\",\"reportTime\":\"2020-11-12\"},{\"deduction\":[\"YLZHT\",\"YLFBHT\",\"NSZHT\",\"NSFBHT\"],\"level\":2,\"orgCode\":[\"JXSB_0106130000\"],\"orgName\":\"江西装备公司\",\"reportTime\":\"2020-11-12\"},{\"deduction\":[\"YLZHT\",\"YLFBHT\",\"NSZHT\",\"NSFBHT\"],\"level\":2,\"orgCode\":[\"SHXZ_0106060000\"],\"orgName\":\"上海装备公司\",\"reportTime\":\"2020-11-12\"},{\"deduction\":[\"YLZHT\",\"YLFBHT\",\"NSZHT\",\"NSFBHT\"],\"level\":2,\"orgCode\":[\"SCJX_0106150000\"],\"orgName\":\"电建透平公司\",\"reportTime\":\"2020-11-12\"},{\"deduction\":[\"YLZHT\",\"YLFBHT\",\"NSZHT\",\"NSFBHT\"],\"level\":2,\"orgCode\":[\"HKSB_0106090000\"],\"orgName\":\"湖北装备公司\",\"reportTime\":\"2020-11-12\"},{\"deduction\":[\"YLZHT\",\"YLFBHT\",\"NSZHT\",\"NSFBHT\"],\"level\":2,\"orgCode\":[\"WHSB_0106080000\"],\"orgName\":\"电建重工公司\",\"reportTime\":\"2020-11-12\"},{\"deduction\":[\"YLZHT\",\"YLFBHT\",\"NSZHT\",\"NSFBHT\"],\"level\":2,\"orgCode\":[\"ZBYJ_0106200000\"],\"orgName\":\"电建装备研究院\",\"reportTime\":\"2020-11-12\"},{\"deduction\":[\"YLZHT\",\"YLFBHT\",\"NSZHT\",\"NSFBHT\"],\"level\":1,\"orgCode\":[\"BJGC_0101020000\",\"JLGC_0101010000\",\"SXGC_0101030000\",\"CHGC_0101040000\",\"SCGC_0101050000\",\"LNGC_0101060000\",\"SCGC_0101070000\",\"HNGC_0101080000\",\"GZGC_0101090000\",\"SCGC_0101100000\",\"HNGC_0101110000\",\"ZJGC_0101120000\",\"TJGC_0101130000\",\"YNGC_0101140000\",\"SXGC_0101150000\",\"FJGC_0101160000\",\"TJJC_0101170000\",\"TJHG_0101260000\"],\"orgName\":\"水电施工\",\"reportTime\":\"2020-11-12\"},{\"deduction\":[\"YLZHT\",\"YLFBHT\",\"NSZHT\",\"NSFBHT\"],\"level\":2,\"orgCode\":[\"BJGC_0101020000\"],\"orgName\":\"电建建筑公司\",\"reportTime\":\"2020-11-12\"},{\"deduction\":[\"YLZHT\",\"YLFBHT\",\"NSZHT\",\"NSFBHT\"],\"level\":2,\"orgCode\":[\"JLGC_0101010000\"],\"orgName\":\"水电一局\",\"reportTime\":\"2020-11-12\"},{\"deduction\":[\"YLZHT\",\"YLFBHT\",\"NSZHT\",\"NSFBHT\"],\"level\":2,\"orgCode\":[\"SXGC_0101030000\"],\"orgName\":\"水电三局\",\"reportTime\":\"2020-11-12\"},{\"deduction\":[\"YLZHT\",\"YLFBHT\",\"NSZHT\",\"NSFBHT\"],\"level\":2,\"orgCode\":[\"CHGC_0101040000\"],\"orgName\":\"水电四局\",\"reportTime\":\"2020-11-12\"},{\"deduction\":[\"YLZHT\",\"YLFBHT\",\"NSZHT\",\"NSFBHT\"],\"level\":2,\"orgCode\":[\"SCGC_0101050000\"],\"orgName\":\"水电五局\",\"reportTime\":\"2020-11-12\"},{\"deduction\":[\"YLZHT\",\"YLFBHT\",\"NSZHT\",\"NSFBHT\"],\"level\":2,\"orgCode\":[\"LNGC_0101060000\"],\"orgName\":\"水电六局\",\"reportTime\":\"2020-11-12\"},{\"deduction\":[\"YLZHT\",\"YLFBHT\",\"NSZHT\",\"NSFBHT\"],\"level\":2,\"orgCode\":[\"SCGC_0101070000\"],\"orgName\":\"水电七局\",\"reportTime\":\"2020-11-12\"},{\"deduction\":[\"YLZHT\",\"YLFBHT\",\"NSZHT\",\"NSFBHT\"],\"level\":2,\"orgCode\":[\"HNGC_0101080000\"],\"orgName\":\"水电八局\",\"reportTime\":\"2020-11-12\"},{\"deduction\":[\"YLZHT\",\"YLFBHT\",\"NSZHT\",\"NSFBHT\"],\"level\":2,\"orgCode\":[\"GZGC_0101090000\"],\"orgName\":\"水电九局\",\"reportTime\":\"2020-11-12\"},{\"deduction\":[\"YLZHT\",\"YLFBHT\",\"NSZHT\",\"NSFBHT\"],\"level\":2,\"orgCode\":[\"SCGC_0101100000\"],\"orgName\":\"水电十局\",\"reportTime\":\"2020-11-12\"},{\"deduction\":[\"YLZHT\",\"YLFBHT\",\"NSZHT\",\"NSFBHT\"],\"level\":2,\"orgCode\":[\"HNGC_0101110000\"],\"orgName\":\"水电十一局\",\"reportTime\":\"2020-11-12\"},{\"deduction\":[\"YLZHT\",\"YLFBHT\",\"NSZHT\",\"NSFBHT\"],\"level\":2,\"orgCode\":[\"ZJGC_0101120000\"],\"orgName\":\"水电十二局\",\"reportTime\":\"2020-11-12\"},{\"deduction\":[\"YLZHT\",\"YLFBHT\",\"NSZHT\",\"NSFBHT\"],\"level\":2,\"orgCode\":[\"TJGC_0101130000\"],\"orgName\":\"电建市政公司\",\"reportTime\":\"2020-11-12\"},{\"deduction\":[\"YLZHT\",\"YLFBHT\",\"NSZHT\",\"NSFBHT\"],\"level\":2,\"orgCode\":[\"YNGC_0101140000\"],\"orgName\":\"水电十四局\",\"reportTime\":\"2020-11-12\"},{\"deduction\":[\"YLZHT\",\"YLFBHT\",\"NSZHT\",\"NSFBHT\"],\"level\":2,\"orgCode\":[\"SXGC_0101150000\"],\"orgName\":\"水电十五局\",\"reportTime\":\"2020-11-12\"},{\"deduction\":[\"YLZHT\",\"YLFBHT\",\"NSZHT\",\"NSFBHT\"],\"level\":2,\"orgCode\":[\"FJGC_0101160000\"],\"orgName\":\"水电十六局\",\"reportTime\":\"2020-11-12\"},{\"deduction\":[\"YLZHT\",\"YLFBHT\",\"NSZHT\",\"NSFBHT\"],\"level\":2,\"orgCode\":[\"TJJC_0101170000\"],\"orgName\":\"水电基础局\",\"reportTime\":\"2020-11-12\"},{\"deduction\":[\"YLZHT\",\"YLFBHT\",\"NSZHT\",\"NSFBHT\"],\"level\":2,\"orgCode\":[\"TJHG_0101260000\"],\"orgName\":\"电建港航公司\",\"reportTime\":\"2020-11-12\"},{\"deduction\":[\"NSFBHT\"],\"level\":1,\"orgCode\":[\"CHIN_160000\",\"CHIN_360000\",\"BJGC_0101020000\",\"YNGW_0102070000\",\"BJGW_0102010000\",\"ZJGW_0102020000\",\"SXGW_0102030000\",\"HNGW_0102040000\",\"SCGW_0102050000\",\"GZGW_0102060000\",\"SCKF_0101210000\",\"BJNY_0101240000\",\"GWJT_0102140000\",\"GSJS_0101270000\",\"BJHT_0101420000\",\"BJLQ_0101200000\",\"BJTL_0101230000\",\"SHJZ_0101330000\",\"BJGJ_0101190000\",\"DJJJ_0109000000\",\"CWGS_0101320000\",\"DJBL_0111000000\",\"BJHK_0101280000\",\"HBYJ_0105010000\",\"SDHD_0105030000\",\"SDYG_0105040000\",\"JLSJ_0104020000\",\"SHSJ_0104030000\",\"FJSJ_0104040000\",\"HNKC_0104060000\",\"JXSJ_0104110000\",\"SCKC_0104070000\",\"QHSJ_0104120000\",\"HBSJ_0104010000\",\"CQJS_0105200000\",\"SDEG_0105050000\",\"SDSG_0105060000\",\"SHJS_0105080000\",\"HBGC_0101500000\",\"HNEJ_0105150000\",\"JXHD_0105160000\",\"JXSD_0105170000\",\"GZYG_0105230000\",\"GZSJ_0104090000\",\"BJFC_0101220000\",\"JLGC_0101010000\",\"SXGC_0101030000\",\"CHGC_0101040000\",\"SCGC_0101050000\",\"LNGC_0101060000\",\"SCGC_0101070000\",\"HNGC_0101080000\",\"GZGC_0101090000\",\"SCGC_0101100000\",\"HNGC_0101110000\",\"ZJGC_0101120000\",\"TJGC_0101130000\",\"YNGC_0101140000\",\"SXGC_0101150000\",\"DJDG_0106210000\",\"CCSB_0106030000\",\"CDJJ_0106170000\",\"HBSB_0106010000\",\"JXSB_0106130000\",\"SHXZ_0106060000\",\"SCJX_0106150000\",\"HKSB_0106090000\",\"WHSB_0106080000\",\"ZBYJ_0106200000\",\"FJGC_0101160000\",\"TJJC_0101170000\",\"TJHG_0101260000\",\"BJZP_0101250000\"],\"orgName\":\"调整数\",\"reportTime\":\"2020-11-12\"}]";
        //数据对象化
        List<NewContractStatisticsTableDco> students = JSON.parseObject(json, new TypeReference<List<NewContractStatisticsTableDco>>() {
        });
        //开始做出处理
        deriveNewContractExcel(students, true);
    }

    //一级层级计数
    static Integer oneSort = 1;
    //二级层级计数
    static Integer twoSort = 1;

    public static void deriveNewContractExcel(List<?> query, Boolean isAll) {
        //一级层级初始化
        oneSort = 1;
        try {
            List<NewContractStatisticsTableDco> copy = new ArrayList<>();
            //进行结果赋值
            query.forEach(i -> {
                NewContractStatisticsTableDco newContractStatisticsTableDco = new NewContractStatisticsTableDco();
                BeanUtils.copyProperties(i, newContractStatisticsTableDco);
                if (newContractStatisticsTableDco.getLevel() == 1) {
                    newContractStatisticsTableDco.setOrgSortId(String.valueOf(oneSort));
                    System.out.println(oneSort);
                    //进行二级层级计数重置
                    twoSort = 1;
                    ++oneSort;
                } else {
                    //不进行成员数值变化
                    Integer temporaryOneSort = oneSort - 1;
                    newContractStatisticsTableDco.setOrgSortId(temporaryOneSort + "." + twoSort);
                    ++twoSort;
                }
                if (isAll) {
                    copy.add(newContractStatisticsTableDco);
                } else {
                    //创建正则表达式
                    Pattern pattern = Pattern.compile("^[1-9]\\d*$");
                    if (pattern.matcher(newContractStatisticsTableDco.getOrgSortId()).matches()) {
                        copy.add(newContractStatisticsTableDco);
                    }
                }
            });
            // 源模版地址
            String templatePath = "/Users/ironegg/Documents/template_新签统计.xls";
            //输出模版地址
            String outPath = "/Users/ironegg/Documents/weijiacutemplate_新签统计.xls";
            //开始根据模版进行填充数据
            ExcelWriter excelWriter = EasyExcel.write(outPath).withTemplate(templatePath).build();
            WriteSheet writeSheet = EasyExcel.writerSheet().build();
            // 直接写入数据
            excelWriter.fill(copy, writeSheet);
            // 写入list之前的数据
            Map<String, Object> map = new HashMap<String, Object>();
            map.put("reportTime", copy.get(0).getReportTime());
            excelWriter.fill(map, writeSheet);
            // 这里是write 别和fill 搞错了
            excelWriter.finish();
            ExcelFontBold(outPath);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }


    public static void ExcelFontBold(String sourceFilePath) {
        //读取源文件流
        InputStream io;
        try {
            io = new FileInputStream(new File(sourceFilePath));
            //读取源文件流转换为对象
            HSSFWorkbook workbook = new HSSFWorkbook(io);
            //读取工作簿对象的表
            HSSFSheet sheet = workbook.getSheetAt(0);
            Integer count = 0;
            Iterator<Row> rowIterator = sheet.rowIterator();
            //固定格式
            HSSFCellStyle fixationCellStyle = sheet.getRow(6).getCell(1).getCellStyle();
            //循环行数据
            while (rowIterator.hasNext()) {
                //获取行数据
                HSSFRow row = sheet.getRow(count);
                //如果行数据不为空
                HSSFCell ongIDCell = null;
                HSSFCell ongNameCell = null;
                if (row != null) {
                    ongIDCell = row.getCell(0);
                    ongNameCell = row.getCell(1);
                    System.out.println(ongIDCell);
                    System.out.println(ongNameCell);
                }
                //创建正则表达式 匹配正数
                Pattern matchingFigure = Pattern.compile("^[1-9]\\d*$");
                //创建正则表达式 匹配小数
                Pattern matchingDecimals = Pattern.compile("^[1-9]\\d*\\.\\d*|0\\.\\d*[1-9]\\d*$");
                if (matchingFigure.matcher(ongIDCell != null ? ongIDCell.toString() : "").matches() ||
                        matchingDecimals.matcher(ongIDCell != null ? ongIDCell.toString() : "").matches()) {
                    row.setHeight((short) 1000);
                }
                if (matchingFigure.matcher(ongIDCell != null ? ongIDCell.toString() : "").matches()) {
                    //加粗企业名称
                    ongNameCell.setCellValue(String.valueOf(ongNameCell));
                    ongIDCell.setCellValue(String.valueOf(ongIDCell));
                    //创建新列格式
                    CellStyle newOngNameCellStyle = workbook.createCellStyle();
                    CellStyle newOngIDCellStyle = workbook.createCellStyle();
                    //克隆旧格式
                    newOngNameCellStyle.cloneStyleFrom(fixationCellStyle);
                    newOngIDCellStyle.cloneStyleFrom(fixationCellStyle);
                    //设置字体
                    newOngNameCellStyle.setFont(createFont(workbook));
                    newOngIDCellStyle.setFont(createFont(workbook));
                    //企业名称列格式赋值
                    ongNameCell.setCellStyle(newOngNameCellStyle);
                    ongIDCell.setCellStyle(newOngNameCellStyle);
                    System.out.println(ongNameCell.getCellStyle());
                    System.out.println(ongIDCell.getCellStyle());
                }
                if (ongIDCell == null) {
                    break;
                }
                count++;
            }
            FileOutputStream fileOutputStream = new FileOutputStream(new File(sourceFilePath));
            workbook.write(fileOutputStream);
            fileOutputStream.close();
            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static HSSFFont createFont(HSSFWorkbook workbook) {
        //创建字体样式
        HSSFFont font = workbook.createFont();
        //设置字体
        font.setFontName("宋体");
        //设置字的大小
        font.setFontHeightInPoints((short) 9);
        //设置粗体
        font.setBold(true);
        return font;
    }
}
