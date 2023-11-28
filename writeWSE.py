def main():
    # Get's WSE from the hdf planFile at the inserted cellNo
    import h5py
    import sys
    import os.path
    filePath = str(sys.argv[1])
    cellNo = int(sys.argv[2])
    wseOutputLocation = str(sys.argv[3])
    with h5py.File(filePath, "r") as f:
        hdfPath = 'Results/Unsteady/Output/Output Blocks/Base Output/Unsteady Time Series/2D Flow Areas/MyRegion/Water Surface'
        WSE = f[hdfPath][()]
        hdfPath = '/Results/Unsteady/Output/Output Blocks/Base Output/Unsteady Time Series/Time Date Stamp'
        time = f[hdfPath][()]
        f.close()
    file = open(wseOutputLocation, "a")
    for x in range(len(WSE)):
        toWrite = time[x].decode('UTF-8') + "," + str(WSE[x][cellNo]) + "\n"
        print(toWrite)
        file.write(toWrite)
    print(os.path.realpath(file.name))
    file.close()

if __name__ == '__main__':
    main()

