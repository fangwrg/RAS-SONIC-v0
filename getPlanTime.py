def main():
    # Get's plantime from the hdf planFile at the inserted cellNo
    import h5py
    import sys
    import os.path
    filePath = str(sys.argv[1])
    cellNo = int(sys.argv[2])
    outputTimePeriodLocation = str(sys.argv[3])
    # filePath = "C:/Users/axp2407/Desktop/Test HDF to ASCII/LakeLivingston.p03.hdf"
    with h5py.File(filePath, "r") as f:
        hdfPath = '/Plan Data/Plan Information/'
        timeWindow = f[hdfPath].attrs["Time Window"].decode()
        shouldBe = "01Mar2017,01:00,01Apr2017,01:00"  # is 01Mar2017 01:00:00 to 01Apr2017 01:00:00
        timeWindow = timeWindow.replace(" to ", ",").replace(":00,",",")[:-3].replace(" ",",")
        f.close()
    file = open(outputTimePeriodLocation, "w")
    file.write(timeWindow)
    print(os.path.realpath(file.name))
    file.close()

if __name__ == '__main__':
    main()
