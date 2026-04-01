from rapidfuzz import fuzz, process

str1 = "Узел измерений\nСбросной газ на УПГ"
str2 = "Узел измерений Сбросной газ на УПГ"

print("token_sort_ratio:", fuzz.token_sort_ratio(str1, str2))
print("token_set_ratio:", fuzz.token_set_ratio(str1, str2))
print("WRatio:", fuzz.WRatio(str1, str2))
print("ratio:", fuzz.ratio(str1, str2))

str1_clean = " ".join(str1.split())
str2_clean = " ".join(str2.split())
print("\nCleaned strings:")
print("token_sort_ratio:", fuzz.token_sort_ratio(str1_clean, str2_clean))
