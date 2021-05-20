from statistics import mean, pstdev, pvariance, stdev, variance
class dataDeal():
    def __init__(self):
        pass

    # 获取众数
    @classmethod
    def modeNums(cls,nums: [int], resultsNum = 3):
        if not isinstance(nums, list):
            print("参数类型错误, 请输入list类型..")
            return False
        else:
            nums.sort(reverse=False)
            startPos = 0
            endPos = startPos + (resultsNum-1)
            if endPos >= len(nums):
                return nums
            elif resultsNum == 0:
                return []

            resultMatch = nums[endPos] - nums[startPos]
            resultPos = startPos
            while endPos < len(nums):
                curResult = nums[endPos] - nums[startPos]
                if curResult < resultMatch:
                    resultPos = startPos
                startPos +=1
                endPos = startPos + (resultsNum-1)

            result = []
            for num in range(resultsNum):
                result.append(nums[resultPos+num])
            return result

    # 获取 平均值 返回list
    @classmethod
    def avgNum(cls, nums: [int]):
        if len(nums) < 1:
            return nums
        return [mean(nums)]

    # 获取 总体标准差 返回list
    @classmethod
    def pstdevNum(cls, nums: [int]):
        if len(nums) < 2:
            return nums
        return [pstdev(nums)]

    # 获取 总体方差 返回list
    @classmethod
    def pvarianceNum(cls, nums: [int]):
        if len(nums) < 2:
            return nums
        return [pvariance(nums)]

    # 获取 样本标准差 返回list
    @classmethod
    def stdevNum(cls, nums: [int]):
        if len(nums) < 2:
            return nums
        return [stdev(nums)]

    # 获取 样本方差 返回list
    @classmethod
    def varianceNum(cls, nums: [int]):
        if len(nums)<2:
            return nums
        return [variance(nums)]

    # 取三个数字，样本标准差 大于0.05就舍弃  小于0.05就保留  取满足要求的平均值最大的三个数
    @classmethod
    def customStdev(cls, nums: [int], resultsNum = 3):
        if len(nums) < 2:
            return False, 0, nums
        elif len(nums) < resultsNum:
            if cls.stdevNum(nums)[0] <= 0.05:
                return False, cls.stdevNum(nums)[0], nums
            else:
                return False, 400, []
        else:
            nums.sort(reverse = False)
            window = [nums[i] for i in range(resultsNum)]
            result = window.copy()
            for item in range(resultsNum,len(nums)):
                window.pop(0)
                window.append(nums[item])

                if cls.stdevNum(window)[0] <= 0.05 and cls.avgNum(window)[0] >= cls.avgNum(result)[0]:
                    result = window.copy()
            if cls.stdevNum(result)[0] <= 0.05:
                return True, cls.stdevNum(result)[0], result
            else:
                return False, cls.stdevNum(result)[0], []





if __name__ == "__main__":
    nums1 = [1.01,1.02,1.03,1.04,1.05,1.06,1.05,1.04,1.07,1.09,1.13]
    res = dataDeal.customStdev(nums1, resultsNum=7)
    print(res)