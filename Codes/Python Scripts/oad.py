exp = Zen.Acquisition.Experiments.GetByName(r'C:\Users\ZSPANIYA\Documents\Carl Zeiss\ZEN\Documents\Experiment Setups\ExperimentDemo.czexp')
img = Zen.Acquisition.Execute(exp)
Zen.Application.Documents.Add(img)

# img = Zen.Application.LoadImage(r'C:\OAD\Input\test.czi', False)
# Zen.Application.Documents.Add(img)
# print(img.Metadata)
#  img.Save(r'C:\temp\testTCPIP.czi')