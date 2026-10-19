from script_zakupki.main_file_zakupki import main, parse_args


if __name__ == "__main__":
    args = parse_args()
    main(
        workbook_path=args.excel,
        files_dir=args.files,
        word_paths=args.word,
        write_enabled=args.write,
    )
