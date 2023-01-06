from os import DirEntry
import re

# finds directory based off of job number
def findDir(jobNumString):
    # dir instantiation
    dir = ""
    if jobNumString == '':
        return "Empty"

    # makes sure jobNum Entry is valid
    if not (jobNumString == ""):
        try:
            jobSArray = re.split('(\\d+)', jobNumString)
            jobNum = int(jobSArray[1])
            jobNumChar = jobSArray[2]
            # finds the first and second digit in the job number
            jobNumD0 = jobNum
            while jobNumD0 >= 10:
                jobNumD0 = jobNumD0 // 10
            jobNumD1 = jobNum
            while jobNumD1 >= 100:
                jobNumD1 = jobNumD1 // 10
            jobNumD1 = jobNumD1 % 10
        except:
            return dir
    else:
        jobNum = int

    if isinstance(jobNum, int) == True and len(jobNumChar) <= 1:
        if len(jobNumString) == 5:
            if jobNumD0 == 1:
                if jobNumD1 == 0:
                    dir = "M:\\10000_10999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                    return dir
                elif jobNumD1 == 1:
                    dir = "M:\\11000_11999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                    return dir
                elif jobNumD1 == 2:
                    dir = "M:\\12000_12999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                    return dir
                elif jobNumD1 == 3:
                    dir = "M:\\13000_13999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                    return dir
                elif jobNumD1 == 4:
                    dir = "M:\\14000_14999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                    return dir
                elif jobNumD1 == 5:
                    dir = "M:\\15000_15999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                    return dir
                elif jobNumD1 == 6:
                    dir = "M:\\16000_16999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                    return dir
                elif jobNumD1 == 7:
                    dir = "M:\\17000_17999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                    return dir
                elif jobNumD1 == 8:
                    dir = "M:\\18000_18999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                    return dir
                elif jobNumD1 == 9:
                    dir = "M:\\19000_19999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                    return dir
                else:
                    return dir
            elif jobNumD0 == 2:
                if jobNumD1 == 0:
                    dir = "M:\\20000_20999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                    return dir
                elif jobNumD1 == 1:
                    dir = "M:\\21000_21999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                    return dir
                elif jobNumD1 == 2:
                    dir = "M:\\22000_22999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                    return dir
                elif jobNumD1 == 3:
                    dir = "M:\\23000_23999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                    return dir
                elif jobNumD1 == 4:
                    dir = "M:\\24000_24999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                    return dir
                elif jobNumD1 == 5:
                    dir = "M:\\25000_25999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                    return dir
                elif jobNumD1 == 6:
                    dir = "M:\\26000_26999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                    return dir
                elif jobNumD1 == 7:
                    dir = "M:\\27000_27999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                    return dir
                elif jobNumD1 == 8:
                    dir = "M:\\28000_28999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                    return dir
                elif jobNumD1 == 9:
                    dir = "M:\\29000_29999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                    return dir
                else:
                    return dir
            elif jobNumD0 == 3:
                if jobNumD1 == 0:
                    dir = "M:\\30000_30999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                    return dir
                elif jobNumD1 == 1:
                    dir = "M:\\31000_31999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                    return dir
                elif jobNumD1 == 2:
                    dir = "M:\\32000_32999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                    return dir
                elif jobNumD1 == 3:
                    dir = "M:\\33000_33999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                    return dir
                elif jobNumD1 == 4:
                    dir = "M:\\34000_34999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                    return dir
                elif jobNumD1 == 5:
                    dir = "M:\\35000_35999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                    return dir
                elif jobNumD1 == 6:
                    dir = "M:\\36000_36999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                    return dir
                elif jobNumD1 == 7:
                    dir = "M:\\37000_37999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                    return dir
                elif jobNumD1 == 8:
                    dir = "M:\\38000_38999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                    return dir
                elif jobNumD1 == 9:
                    dir = "M:\\39000_39999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                    return dir
                else:
                    return dir
            elif jobNumD0 == 4:
                if jobNumD1 == 0:
                    dir = "M:\\40000_40999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                    return dir
                elif jobNumD1 == 1:
                    dir = "M:\\41000_41999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                    return dir
                elif jobNumD1 == 2:
                    dir = "M:\\42000_42999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                    return dir
                elif jobNumD1 == 3:
                    dir = "M:\\43000_43999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                    return dir
                elif jobNumD1 == 4:
                    dir = "M:\\44000_44999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                    return dir
                elif jobNumD1 == 5:
                    dir = "M:\\45000_45999\\" + jobNumString + "\\Electrical\\" + jobNumString+ " Drawings"
                    return dir
                elif jobNumD1 == 6:
                    dir = "M:\\46000_46999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                    return dir
                elif jobNumD1 == 7:
                    dir = "M:\\47000_47999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                    return dir
                elif jobNumD1 == 8:
                    dir = "M:\\48000_48999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                    return dir
                elif jobNumD1 == 9:
                    dir = "M:\\49000_49999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                    return dir
                else:
                    return dir
            elif jobNumD0 == 5:
                if jobNumD1 == 0:
                    dir = "M:\\50000_50999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                    return dir
                elif jobNumD1 == 1:
                    dir = "M:\\51000_51999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                    return dir
                elif jobNumD1 == 2:
                    dir = "M:\\52000_52999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                    return dir
                elif jobNumD1 == 3:
                    dir = "M:\\53000_53999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                    return dir
                elif jobNumD1 == 4:
                    dir = "M:\\54000_54999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                    return dir
                elif jobNumD1 == 5:
                    dir = "M:\\55000_55999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                    return dir
                elif jobNumD1 == 6:
                    dir = "M:\\56000_56999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                    return dir
                elif jobNumD1 == 7:
                    dir = "M:\\57000_57999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                    return dir
                elif jobNumD1 == 8:
                    dir = "M:\\58000_58999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                    return dir
                elif jobNumD1 == 9:
                    dir = "M:\\59000_59999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                    return dir
                else:
                    return dir
            elif jobNumD0 == 6:
                if jobNumD1 == 0:
                    dir = "M:\\60000_60999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                    return dir
                elif jobNumD1 == 1:
                    dir = "M:\\61000_61999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                    return dir
                elif jobNumD1 == 2:
                    dir = "M:\\62000_62999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                    return dir
                elif jobNumD1 == 3:
                    dir = "M:\\63000_63999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                    return dir
                elif jobNumD1 == 4:
                    dir = "M:\\64000_64999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                    return dir
                elif jobNumD1 == 5:
                    dir = "M:\\65000_65999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                    return dir
                elif jobNumD1 == 6:
                    dir = "M:\\66000_66999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                    return dir
                elif jobNumD1 == 7:
                    dir = "M:\\67000_67999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                    return dir
                elif jobNumD1 == 8:
                    dir = "M:\\68000_68999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                    return dir
                elif jobNumD1 == 9:
                    dir = "M:\\69000_69999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                    return dir
                else:
                    return dir
            elif jobNumD0 == 7:
                if jobNumD1 == 0:
                    dir = "M:\\70000_70999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                    return dir
                elif jobNumD1 == 1:
                    dir = "M:\\71000_71999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                    return dir
                elif jobNumD1 == 2:
                    dir = "M:\\72000_72999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                    return dir
                elif jobNumD1 == 3:
                    dir = "M:\\73000_73999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                    return dir
                elif jobNumD1 == 4:
                    dir = "M:\\74000_74999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                    return dir
                elif jobNumD1 == 5:
                    dir = "M:\\75000_75999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                    return dir
                elif jobNumD1 == 6:
                    dir = "M:\\76000_76999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                    return dir
                elif jobNumD1 == 7:
                    dir = "M:\\77000_77999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                    return dir
                elif jobNumD1 == 8:
                    dir = "M:\\78000_78999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                    return dir
                elif jobNumD1 == 9:
                    dir = "M:\\79000_79999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                    return dir
                else:
                    return dir
            elif jobNumD0 == 8:
                if jobNumD1 == 0:
                    dir = "M:\\80000_80999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                    return dir
                elif jobNumD1 == 1:
                    dir = "M:\\81000_81999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                    return dir
                elif jobNumD1 == 2:
                    dir = "M:\\82000_82999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                    return dir
                elif jobNumD1 == 3:
                    dir = "M:\\83000_83999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                    return dir
                elif jobNumD1 == 4:
                    dir = "M:\\84000_84999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                    return dir
                elif jobNumD1 == 5:
                    dir = "M:\\85000_85999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                    return dir
                elif jobNumD1 == 6:
                    dir = "M:\\86000_86999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                    return dir
                elif jobNumD1 == 7:
                    dir = "M:\\87000_87999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                    return dir
                elif jobNumD1 == 8:
                    dir = "M:\\88000_88999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                    return dir
                elif jobNumD1 == 9:
                    dir = "M:\\89000_89999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                    return dir
                else:
                    return dir
            elif jobNumD0 == 9:
                if jobNumD1 == 0:
                    dir = "M:\\90000_90999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                    return dir
                elif jobNumD1 == 1:
                    dir = "M:\\91000_91999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                    return dir
                elif jobNumD1 == 2:
                    dir = "M:\\92000_92999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                    return dir
                elif jobNumD1 == 3:
                    dir = "M:\\93000_93999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                    return dir
                elif jobNumD1 == 4:
                    dir = "M:\\94000_94999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                    return dir
                elif jobNumD1 == 5:
                    dir = "M:\\95000_95999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                    return dir
                elif jobNumD1 == 6:
                    dir = "M:\\96000_96999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                    return dir
                elif jobNumD1 == 7:
                    dir = "M:\\97000_97999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                    return dir
                elif jobNumD1 == 8:
                    dir = "M:\\98000_98999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                    return dir
                elif jobNumD1 == 9:
                    dir = "M:\\99000_99999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                    return dir
                else:
                    return dir
            else:
                return dir
        elif len(jobNumString) == 4:
            if jobNumD0 < 4:
                dir = "M:\\0000_9999\\0000_3999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                return dir
            elif jobNumD1 == 4:
                dir = "M:\\0000_9999\\4000_4999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                return dir
            elif jobNumD1 == 5:
                dir = "M:\\0000_9999\\5000_5999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                return dir
            elif jobNumD1 == 6:
                dir = "M:\\0000_9999\\6000_6999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                return dir
            elif jobNumD1 == 7:
                dir = "M:\\0000_9999\\7000_7999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                return dir
            elif jobNumD1 == 8:
                dir = "M:\\0000_9999\\8000_8999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                return dir
            elif jobNumD1 == 9:
                dir = "M:\\0000_9999\\9000_9999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                return dir
            else:
                return dir
        elif len(jobNumString) == 3 or len(jobNumString) == 2 or len(jobNumString) == 1:
            return "Short"
        elif len(jobNumString) == 0 or (jobNum == '' and jobNumChar == ''):
            return "Empty"
        elif len(jobNumString) == 6:
            if len(jobNumChar) > 1 or len(jobSArray[1]) > 5:
                return "Long"
            else:
                if jobNumD0 == 1:
                    if jobNumD1 == 0:
                        dir = "M:\\10000_10999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                        return dir
                    elif jobNumD1 == 1:
                        dir = "M:\\11000_11999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                        return dir
                    elif jobNumD1 == 2:
                        dir = "M:\\12000_12999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                        return dir
                    elif jobNumD1 == 3:
                        dir = "M:\\13000_13999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                        return dir
                    elif jobNumD1 == 4:
                        dir = "M:\\14000_14999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                        return dir
                    elif jobNumD1 == 5:
                        dir = "M:\\15000_15999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                        return dir
                    elif jobNumD1 == 6:
                        dir = "M:\\16000_16999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                        return dir
                    elif jobNumD1 == 7:
                        dir = "M:\\17000_17999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                        return dir
                    elif jobNumD1 == 8:
                        dir = "M:\\18000_18999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                        return dir
                    elif jobNumD1 == 9:
                        dir = "M:\\19000_19999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                        return dir
                    else:
                        return dir
                elif jobNumD0 == 2:
                    if jobNumD1 == 0:
                        dir = "M:\\20000_20999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                        return dir
                    elif jobNumD1 == 1:
                        dir = "M:\\21000_21999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                        return dir
                    elif jobNumD1 == 2:
                        dir = "M:\\22000_22999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                        return dir
                    elif jobNumD1 == 3:
                        dir = "M:\\23000_23999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                        return dir
                    elif jobNumD1 == 4:
                        dir = "M:\\24000_24999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                        return dir
                    elif jobNumD1 == 5:
                        dir = "M:\\25000_25999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                        return dir
                    elif jobNumD1 == 6:
                        dir = "M:\\26000_26999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                        return dir
                    elif jobNumD1 == 7:
                        dir = "M:\\27000_27999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                        return dir
                    elif jobNumD1 == 8:
                        dir = "M:\\28000_28999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                        return dir
                    elif jobNumD1 == 9:
                        dir = "M:\\29000_29999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                        return dir
                    else:
                        return dir
                elif jobNumD0 == 3:
                    if jobNumD1 == 0:
                        dir = "M:\\30000_30999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                        return dir
                    elif jobNumD1 == 1:
                        dir = "M:\\31000_31999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                        return dir
                    elif jobNumD1 == 2:
                        dir = "M:\\32000_32999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                        return dir
                    elif jobNumD1 == 3:
                        dir = "M:\\33000_33999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                        return dir
                    elif jobNumD1 == 4:
                        dir = "M:\\34000_34999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                        return dir
                    elif jobNumD1 == 5:
                        dir = "M:\\35000_35999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                        return dir
                    elif jobNumD1 == 6:
                        dir = "M:\\36000_36999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                        return dir
                    elif jobNumD1 == 7:
                        dir = "M:\\37000_37999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                        return dir
                    elif jobNumD1 == 8:
                        dir = "M:\\38000_38999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                        return dir
                    elif jobNumD1 == 9:
                        dir = "M:\\39000_39999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                        return dir
                    else:
                        return dir
                elif jobNumD0 == 4:
                    if jobNumD1 == 0:
                        dir = "M:\\40000_40999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                        return dir
                    elif jobNumD1 == 1:
                        dir = "M:\\41000_41999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                        return dir
                    elif jobNumD1 == 2:
                        dir = "M:\\42000_42999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                        return dir
                    elif jobNumD1 == 3:
                        dir = "M:\\43000_43999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                        return dir
                    elif jobNumD1 == 4:
                        dir = "M:\\44000_44999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                        return dir
                    elif jobNumD1 == 5:
                        dir = "M:\\45000_45999\\" + jobNumString + "\\Electrical\\" + jobNumString+ " Drawings"
                        return dir
                    elif jobNumD1 == 6:
                        dir = "M:\\46000_46999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                        return dir
                    elif jobNumD1 == 7:
                        dir = "M:\\47000_47999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                        return dir
                    elif jobNumD1 == 8:
                        dir = "M:\\48000_48999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                        return dir
                    elif jobNumD1 == 9:
                        dir = "M:\\49000_49999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                        return dir
                    else:
                        return dir
                elif jobNumD0 == 5:
                    if jobNumD1 == 0:
                        dir = "M:\\50000_50999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                        return dir
                    elif jobNumD1 == 1:
                        dir = "M:\\51000_51999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                        return dir
                    elif jobNumD1 == 2:
                        dir = "M:\\52000_52999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                        return dir
                    elif jobNumD1 == 3:
                        dir = "M:\\53000_53999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                        return dir
                    elif jobNumD1 == 4:
                        dir = "M:\\54000_54999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                        return dir
                    elif jobNumD1 == 5:
                        dir = "M:\\55000_55999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                        return dir
                    elif jobNumD1 == 6:
                        dir = "M:\\56000_56999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                        return dir
                    elif jobNumD1 == 7:
                        dir = "M:\\57000_57999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                        return dir
                    elif jobNumD1 == 8:
                        dir = "M:\\58000_58999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                        return dir
                    elif jobNumD1 == 9:
                        dir = "M:\\59000_59999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                        return dir
                    else:
                        return dir
                elif jobNumD0 == 6:
                    if jobNumD1 == 0:
                        dir = "M:\\60000_60999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                        return dir
                    elif jobNumD1 == 1:
                        dir = "M:\\61000_61999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                        return dir
                    elif jobNumD1 == 2:
                        dir = "M:\\62000_62999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                        return dir
                    elif jobNumD1 == 3:
                        dir = "M:\\63000_63999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                        return dir
                    elif jobNumD1 == 4:
                        dir = "M:\\64000_64999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                        return dir
                    elif jobNumD1 == 5:
                        dir = "M:\\65000_65999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                        return dir
                    elif jobNumD1 == 6:
                        dir = "M:\\66000_66999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                        return dir
                    elif jobNumD1 == 7:
                        dir = "M:\\67000_67999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                        return dir
                    elif jobNumD1 == 8:
                        dir = "M:\\68000_68999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                        return dir
                    elif jobNumD1 == 9:
                        dir = "M:\\69000_69999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                        return dir
                    else:
                        return dir
                elif jobNumD0 == 7:
                    if jobNumD1 == 0:
                        dir = "M:\\70000_70999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                        return dir
                    elif jobNumD1 == 1:
                        dir = "M:\\71000_71999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                        return dir
                    elif jobNumD1 == 2:
                        dir = "M:\\72000_72999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                        return dir
                    elif jobNumD1 == 3:
                        dir = "M:\\73000_73999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                        return dir
                    elif jobNumD1 == 4:
                        dir = "M:\\74000_74999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                        return dir
                    elif jobNumD1 == 5:
                        dir = "M:\\75000_75999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                        return dir
                    elif jobNumD1 == 6:
                        dir = "M:\\76000_76999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                        return dir
                    elif jobNumD1 == 7:
                        dir = "M:\\77000_77999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                        return dir
                    elif jobNumD1 == 8:
                        dir = "M:\\78000_78999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                        return dir
                    elif jobNumD1 == 9:
                        dir = "M:\\79000_79999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                        return dir
                    else:
                        return dir
                elif jobNumD0 == 8:
                    if jobNumD1 == 0:
                        dir = "M:\\80000_80999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                        return dir
                    elif jobNumD1 == 1:
                        dir = "M:\\81000_81999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                        return dir
                    elif jobNumD1 == 2:
                        dir = "M:\\82000_82999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                        return dir
                    elif jobNumD1 == 3:
                        dir = "M:\\83000_83999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                        return dir
                    elif jobNumD1 == 4:
                        dir = "M:\\84000_84999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                        return dir
                    elif jobNumD1 == 5:
                        dir = "M:\\85000_85999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                        return dir
                    elif jobNumD1 == 6:
                        dir = "M:\\86000_86999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                        return dir
                    elif jobNumD1 == 7:
                        dir = "M:\\87000_87999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                        return dir
                    elif jobNumD1 == 8:
                        dir = "M:\\88000_88999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                        return dir
                    elif jobNumD1 == 9:
                        dir = "M:\\89000_89999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                        return dir
                    else:
                        return dir
                elif jobNumD0 == 9:
                    if jobNumD1 == 0:
                        dir = "M:\\90000_90999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                        return dir
                    elif jobNumD1 == 1:
                        dir = "M:\\91000_91999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                        return dir
                    elif jobNumD1 == 2:
                        dir = "M:\\92000_92999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                        return dir
                    elif jobNumD1 == 3:
                        dir = "M:\\93000_93999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                        return dir
                    elif jobNumD1 == 4:
                        dir = "M:\\94000_94999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                        return dir
                    elif jobNumD1 == 5:
                        dir = "M:\\95000_95999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                        return dir
                    elif jobNumD1 == 6:
                        dir = "M:\\96000_96999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                        return dir
                    elif jobNumD1 == 7:
                        dir = "M:\\97000_97999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                        return dir
                    elif jobNumD1 == 8:
                        dir = "M:\\98000_98999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                        return dir
                    elif jobNumD1 == 9:
                        dir = "M:\\99000_99999\\" + jobNumString + "\\Electrical\\" + jobNumString + " Drawings"
                        return dir
                    else:
                        return dir
                else:
                    return dir
        else:
            return "Long"
    else: 
        return dir