import math
class Service:
    # Constructor

   # MonthlyAmount = 0.0
    #AnnualAmountYear1 = 0.0
    #AnnualAmountYear2 = 0.0
    #AnnualAmountYear3 = 0.0
    def __init__(self,serviceName, weeklyHours, billRate, yearlyHolidayHours,inflationRate,annualAmountYear1,annualAmountYear2,annualAmountYear3,monthlyAmount):
        self.serviceName = serviceName
        self.weeklyHours = weeklyHours
        self.billRate = billRate
        self.yearlyHolidayHours = yearlyHolidayHours
        self.inflationRate = inflationRate/100
        holidayRate = billRate * 1.5
        holidayEffect = (holidayRate - billRate) * self.yearlyHolidayHours
        self.monthlyAmount = math.floor(float((weeklyHours * 52 * billRate) / 12) + holidayEffect)
        self.annualAmountYear1 = math.floor(float(self.monthlyAmount * 12))
        self.annualAmountYear2 = math.floor(self.annualAmountYear1 * (1 + self.inflationRate))
        self.annualAmountYear3 = math.floor(self.annualAmountYear2 * (1 + self.inflationRate))


    #toString
    def __str__(self):
        return f"({self.serviceName})({self.weeklyHours})({self.billRate})({self.monthlyAmount})({self.annualAmountYear1})"

