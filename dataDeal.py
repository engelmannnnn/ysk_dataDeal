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
        return [mean(nums)]

    # 获取 总体标准差 返回list
    @classmethod
    def pstdevNum(cls, nums: [int]):
        return [pstdev(nums)]

    # 获取 总体方差 返回list
    @classmethod
    def pvarianceNum(cls, nums: [int]):
        return [pvariance(nums)]

    # 获取 样本标准差 返回list
    @classmethod
    def stdevNum(cls, nums: [int]):
        return [stdev(nums)]

    # 获取 样本方差 返回list
    @classmethod
    def varianceNum(cls, nums: [int]):
        return [variance(nums)]






if __name__ == "__main__":
    res = dataDeal.modeNums([1.5121263,1.487,1.545,1.403,1.435])
    print(res)