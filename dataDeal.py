class dataDeal():
    def __init__(self):
        pass

    @classmethod
    # 获取众数
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


if __name__ == "__main__":
    res = dataDeal.modeNums([1.5121263,1.487,1.545,1.403,1.435])
    print(res)