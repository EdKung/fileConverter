import exp2xls
import exp2xlsx

def main():
    src_folder = './input'
    output_folder = './output'

    # exp2xls.run(src_folder, output_folder)
    exp2xlsx.run(src_folder, output_folder)

if __name__ == "__main__":
    main()
