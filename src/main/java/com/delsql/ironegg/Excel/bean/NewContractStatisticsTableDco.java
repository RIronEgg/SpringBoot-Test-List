package com.delsql.ironegg.Excel.bean;

import lombok.AccessLevel;
import lombok.Data;
import lombok.Getter;

import java.math.BigDecimal;
import java.util.List;

/**
 * @author RenZiHan
 * @version 1.0
 * @date 2020/11/17 4:00 下午
 * @purpose Excel导出Bean
 */
@Data
public class NewContractStatisticsTableDco {
    private Integer level;
    private List<String> orgCode;
    private String industryTypeCode;
    private String reportTime;
    private List<String> deduction;

    private String orgSortId;

    private String orgName;

    @Getter(AccessLevel.NONE)
    private BigDecimal lastYearData;

    @Getter(AccessLevel.NONE)
    private BigDecimal lastYearSamePeriodComplete;

    @Getter(AccessLevel.NONE)
    private BigDecimal lastPeriodComplete;

    @Getter(AccessLevel.NONE)
    private BigDecimal currentPeriodComplete;

    private BigDecimal yearOnYearGrowth;
    private BigDecimal chainGrowth;

    private BigDecimal annualPlan;

    @Getter(AccessLevel.NONE)
    private BigDecimal lastPeriodYearGrandTotalComplete;

    @Getter(AccessLevel.NONE)
    private BigDecimal yearGrandTotalComplete;

    private BigDecimal completePlan;

    //===========================国内========================================

    @Getter(AccessLevel.NONE)
    //("国内项目")
    private BigDecimal domesticProject;

    @Getter(AccessLevel.NONE)
    //("能源电力业务-国内")
    private BigDecimal dcEnergyPowerBusiness;

    @Getter(AccessLevel.NONE)
    //("水利-国内")
    private BigDecimal dcWaterConservancy;

    @Getter(AccessLevel.NONE)
    //("水电-国内")
    private BigDecimal dcHydropower;

    @Getter(AccessLevel.NONE)
    //("火电-国内")
    private BigDecimal dcThermalPower;

    @Getter(AccessLevel.NONE)
    //("核电-国内")
    private BigDecimal dcNuclearPower;

    @Getter(AccessLevel.NONE)
    //("风电-国内")
    private BigDecimal dcWindPower;

    @Getter(AccessLevel.NONE)
    //("太阳能发电-国内")
    private BigDecimal dcSolarEnergyGeneration;

    @Getter(AccessLevel.NONE)
    //("电网-国内")
    private BigDecimal dcPowerGrid;

    @Getter(AccessLevel.NONE)
    //("能源电力业务其他-国内")
    private BigDecimal dcEnergyPowerBusinessOther;

    @Getter(AccessLevel.NONE)
    //("水资源与环境业务-国内")
    private BigDecimal dcWaterEnvironmentBusiness;

    @Getter(AccessLevel.NONE)
    //("水务工程-国内")
    private BigDecimal dcWaterWorks;

    @Getter(AccessLevel.NONE)
    //("水环境治理和水生态修复工程-国内")
    private BigDecimal dcWaterGovernanceRepair;

    @Getter(AccessLevel.NONE)
    //("工业环保工程（非水）-国内")
    private BigDecimal dcIndustrialEnvironmental;

    @Getter(AccessLevel.NONE)
    //("基础设施业务-国内")
    private BigDecimal dcInfrastructureBusiness;

    @Getter(AccessLevel.NONE)
    //("房屋建筑-国内")
    private BigDecimal dcConstruction;

    @Getter(AccessLevel.NONE)
    //("铁路-国内")
    private BigDecimal dcRailway;

    @Getter(AccessLevel.NONE)
    //("城市轨道交通-国内")
    private BigDecimal dcUrbanRailTransit;

    @Getter(AccessLevel.NONE)
    //("公路-国内")
    private BigDecimal dcHighway;

    @Getter(AccessLevel.NONE)
    //("市政-国内")
    private BigDecimal dcMunicipal;

    @Getter(AccessLevel.NONE)
    //("机场-国内")
    private BigDecimal dcAirport;

    @Getter(AccessLevel.NONE)
    //("港口与航道-国内")
    private BigDecimal dcPortsWaterways;

    @Getter(AccessLevel.NONE)
    //("基础设施业务其他-国内")
    private BigDecimal dcInfrastructureBusinessOther;

    @Getter(AccessLevel.NONE)
    //("房地产业务-国内")
    private BigDecimal dcRealEstateBusiness;

    @Getter(AccessLevel.NONE)
    //("其他业务-国内")
    private BigDecimal dcOtherBusiness;

    //===========================国外========================================
    @Getter(AccessLevel.NONE)
    //("国外项目")
    private BigDecimal foreignProject;

    @Getter(AccessLevel.NONE)
    //("能源电力业务-国外")
    private BigDecimal fnEnergyPowerBusiness;

    @Getter(AccessLevel.NONE)
    //("水利-国外")
    private BigDecimal fnWaterConservancy;

    @Getter(AccessLevel.NONE)
    //("水电-国外")
    private BigDecimal fnHydropower;

    @Getter(AccessLevel.NONE)
    //("火电-国外")
    private BigDecimal fnThermalPower;

    @Getter(AccessLevel.NONE)
    //("核电-国外")
    private BigDecimal fnNuclearPower;

    @Getter(AccessLevel.NONE)
    //("风电-国外")
    private BigDecimal fnWindPower;

    @Getter(AccessLevel.NONE)
    //("太阳能发电-国外")
    private BigDecimal fnSolarEnergyGeneration;

    @Getter(AccessLevel.NONE)
    //("电网-国外")
    private BigDecimal fnPowerGrid;

    @Getter(AccessLevel.NONE)
    //("能源电力业务其他-国外")
    private BigDecimal fnEnergyPowerBusinessOther;

    @Getter(AccessLevel.NONE)
    //("水资源与环境业务-国外")
    private BigDecimal fnWaterEnvironmentBusiness;

    @Getter(AccessLevel.NONE)
    //("水务工程-国外")
    private BigDecimal fnWaterWorks;

    @Getter(AccessLevel.NONE)
    //("水环境治理和水生态修复工程-国外")
    private BigDecimal fnWaterGovernanceRepair;

    @Getter(AccessLevel.NONE)
    //("工业环保工程（非水）-国外")
    private BigDecimal fnIndustrialEnvironmental;

    @Getter(AccessLevel.NONE)
    //("基础设施业务-国外")
    private BigDecimal fnInfrastructureBusiness;

    @Getter(AccessLevel.NONE)
    //("房屋建筑-国外")
    private BigDecimal fnConstruction;

    @Getter(AccessLevel.NONE)
    //("铁路-国外")
    private BigDecimal fnRailway;

    @Getter(AccessLevel.NONE)
    //("城市轨道交通-国外")
    private BigDecimal fnUrbanRailTransit;

    @Getter(AccessLevel.NONE)
    //("公路-国外")
    private BigDecimal fnHighway;

    @Getter(AccessLevel.NONE)
    //("市政-国外")
    private BigDecimal fnMunicipal;

    @Getter(AccessLevel.NONE)
    //("机场-国外")
    private BigDecimal fnAirport;

    @Getter(AccessLevel.NONE)
    //("港口与航道-国外")
    private BigDecimal fnPortsWaterways;

    @Getter(AccessLevel.NONE)
    //("基础设施业务其他-国外")
    private BigDecimal fnInfrastructureBusinessOther;

    @Getter(AccessLevel.NONE)
    //("房地产业务-国外")
    private BigDecimal fnRealEstateBusiness;

    @Getter(AccessLevel.NONE)
    //("其他业务-国外")
    private BigDecimal fnOtherBusiness;
}
