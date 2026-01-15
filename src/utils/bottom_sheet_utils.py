import flet as ft


class BottomSheetTheme:
    BG = ft.Colors.GREY_900
    TEXT = ft.Colors.WHITE
    HEADER_BG = ft.Colors.GREY_900
    DIVIDER = ft.Colors.GREY_800


class BottomSheetManager:
    """
    Bottom Sheet Manager for consistent bottom sheet handling across the app
    """

    @staticmethod
    def show_bottom_sheet(
        page: ft.Page,
        title: str,
        content: ft.Control,
        icon: str = None,
        icon_color: str = None,
        on_dismiss=None,
        show_close_button: bool = True,
    ):
        """
        Show a bottom sheet with consistent styling
        
        Args:
            page: Flet page object
            title: Title text for the bottom sheet
            content: Content control to display
            icon: Optional icon to show next to title
            icon_color: Optional color for the icon
            on_dismiss: Optional callback when bottom sheet is dismissed
            show_close_button: Whether to show close button in header
        """
        def close_bs(e):
            bs.open = False
            bs.update()
            if on_dismiss:
                on_dismiss(e)

        # Build header controls
        header_controls = []
        
        if icon:
            header_controls.append(
                ft.Icon(icon, color=icon_color or ft.Colors.BLUE_400, size=28)
            )
        
        header_controls.append(
            ft.Text(title, weight=ft.FontWeight.BOLD, size=20)
        )
        
        header_controls.append(ft.Container(expand=True))
        
        if show_close_button:
            header_controls.append(
                ft.IconButton(
                    icon=ft.Icons.CLOSE,
                    on_click=close_bs,
                    icon_color=ft.Colors.GREY_400,
                )
            )

        bs = ft.BottomSheet(
            content=ft.Container(
                content=ft.Column(
                    controls=[
                        # Header
                        ft.Container(
                            content=ft.Row(
                                controls=header_controls,
                                alignment=ft.MainAxisAlignment.START,
                            ),
                            padding=ft.padding.only(left=20, right=20, top=20, bottom=10),
                        ),
                        ft.Divider(height=1, color=BottomSheetTheme.DIVIDER),
                        # Content
                        ft.Container(
                            content=content,
                            padding=ft.padding.all(20),
                        ),
                    ],
                    tight=True,
                ),
                bgcolor=BottomSheetTheme.BG,
            ),
            open=True,
            on_dismiss=close_bs if on_dismiss else None,
        )

        page.overlay.append(bs)
        page.update()
        
        return bs

    @staticmethod
    def show_options_bottom_sheet(
        page: ft.Page,
        title: str,
        options: list,
        icon: str = None,
        icon_color: str = None,
        description: str = None,
    ):
        """
        Show a bottom sheet with multiple option cards
        
        Args:
            page: Flet page object
            title: Title text for the bottom sheet
            options: List of option dictionaries with keys:
                - text: Main text for the option
                - subtext: Secondary text (optional)
                - icon: Icon for the option
                - color: Background color for the card
                - on_click: Callback function when option is clicked
            icon: Optional icon to show next to title
            icon_color: Optional color for the icon
            description: Optional description text below title
        """
        # Container to hold the bottom sheet reference
        bs_container = {}
        
        def close_bs(e):
            if "bs" in bs_container:
                bs_container["bs"].open = False
                bs_container["bs"].update()

        # Build option cards
        option_cards = []
        for option in options:
            card_controls = [
                ft.Icon(
                    option.get("icon", ft.Icons.CIRCLE),
                    size=40,
                    color=ft.Colors.WHITE
                ),
                ft.Text(
                    option.get("text", ""),
                    size=15,
                    weight=ft.FontWeight.W_600,
                    text_align=ft.TextAlign.CENTER
                ),
            ]
            
            if option.get("subtext"):
                card_controls.append(
                    ft.Text(
                        option["subtext"],
                        size=11,
                        color=ft.Colors.GREY_400,
                        text_align=ft.TextAlign.CENTER
                    )
                )

            def make_click_handler(callback):
                def handler(e):
                    close_bs(e)
                    if callback:
                        callback(e)
                return handler

            card = ft.Card(
                content=ft.Container(
                    content=ft.Column(
                        controls=card_controls,
                        alignment=ft.MainAxisAlignment.CENTER,
                        horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                        spacing=8,
                        tight=True,
                    ),
                    padding=15,
                    alignment=ft.alignment.center,
                    bgcolor=option.get("color", ft.Colors.BLUE_700),
                    border_radius=15,
                    ink=True,
                    on_click=make_click_handler(option.get("on_click")),
                    width=150,
                    height=130,
                ),
                elevation=8,
            )
            option_cards.append(card)

        # Build content
        content_controls = []
        
        if description:
            content_controls.append(
                ft.Text(description, size=14, color=ft.Colors.GREY_400)
            )
            content_controls.append(ft.Container(height=15))
        
        content_controls.append(
            ft.Row(
                controls=option_cards,
                alignment=ft.MainAxisAlignment.CENTER,
                spacing=20,
                wrap=True,
            )
        )

        content = ft.Column(
            controls=content_controls,
            horizontal_alignment=ft.CrossAxisAlignment.CENTER,
        )

        # Build header controls
        header_controls = []
        
        if icon:
            header_controls.append(
                ft.Icon(icon, color=icon_color or ft.Colors.BLUE_400, size=28)
            )
        
        header_controls.append(
            ft.Text(title, weight=ft.FontWeight.BOLD, size=20)
        )
        
        header_controls.append(ft.Container(expand=True))
        
        header_controls.append(
            ft.IconButton(
                icon=ft.Icons.CLOSE,
                on_click=close_bs,
                icon_color=ft.Colors.GREY_400,
            )
        )

        bs = ft.BottomSheet(
            content=ft.Container(
                content=ft.Column(
                    controls=[
                        # Header
                        ft.Container(
                            content=ft.Row(
                                controls=header_controls,
                                alignment=ft.MainAxisAlignment.START,
                            ),
                            padding=ft.padding.only(left=20, right=20, top=20, bottom=10),
                        ),
                        ft.Divider(height=1, color=BottomSheetTheme.DIVIDER),
                        # Content
                        ft.Container(
                            content=content,
                            padding=ft.padding.all(20),
                        ),
                    ],
                    tight=True,
                ),
                bgcolor=BottomSheetTheme.BG,
            ),
            open=True,
            on_dismiss=close_bs,
        )

        bs_container["bs"] = bs
        page.overlay.append(bs)
        page.update()
        
        return bs

    @staticmethod
    def close_bottom_sheet(bs: ft.BottomSheet):
        """Close a bottom sheet"""
        if bs:
            bs.open = False
            bs.update()

    @staticmethod
    def show_success_bottom_sheet(
        page: ft.Page,
        message: str,
        filepath: str = None,
        title: str = "تم الحفظ بنجاح",
        on_open_file=None,
        on_open_folder=None,
    ):
        """
        Show a success bottom sheet with optional file actions
        
        Args:
            page: Flet page object
            message: Success message to display
            filepath: Optional file path for open actions
            title: Title text for the bottom sheet
            on_open_file: Optional callback for opening file
            on_open_folder: Optional callback for opening folder
        """
        import os
        
        def close_bs(e):
            bs.open = False
            bs.update()
        
        def open_file(e):
            close_bs(e)
            if on_open_file:
                on_open_file(e)
            elif filepath:
                try:
                    os.startfile(filepath)
                except Exception:
                    pass
        
        def open_folder(e):
            close_bs(e)
            if on_open_folder:
                on_open_folder(e)
            elif filepath:
                try:
                    folder = os.path.dirname(filepath)
                    os.startfile(folder)
                except Exception:
                    pass
        
        # Build content
        content_controls = [
            ft.Row(
                controls=[
                    ft.Icon(ft.Icons.CHECK_CIRCLE, color=ft.Colors.GREEN_400, size=40),
                    ft.Container(width=15),
                    ft.Column(
                        controls=[
                            ft.Text(message, size=16, rtl=True, weight=ft.FontWeight.W_500),
                        ],
                        spacing=5,
                    ),
                ],
                alignment=ft.MainAxisAlignment.START,
            ),
        ]
        
        # Add file info if filepath provided
        if filepath:
            content_controls.append(ft.Container(height=15))
            content_controls.append(
                ft.Container(
                    content=ft.Text(
                        os.path.basename(filepath),
                        size=13,
                        color=ft.Colors.BLUE_200,
                        weight=ft.FontWeight.W_500,
                        rtl=True,
                    ),
                    bgcolor=ft.Colors.BLUE_GREY_800,
                    padding=10,
                    border_radius=8,
                )
            )
            
            # Add action buttons
            content_controls.append(ft.Container(height=20))
            content_controls.append(
                ft.Row(
                    controls=[
                        ft.ElevatedButton(
                            "فتح الملف",
                            icon=ft.Icons.FILE_OPEN,
                            on_click=open_file,
                            bgcolor=ft.Colors.GREEN_700,
                            color=ft.Colors.WHITE,
                        ),
                        ft.ElevatedButton(
                            "فتح المجلد",
                            icon=ft.Icons.FOLDER_OPEN,
                            on_click=open_folder,
                            bgcolor=ft.Colors.BLUE_700,
                            color=ft.Colors.WHITE,
                        ),
                        ft.TextButton(
                            "إغلاق",
                            on_click=close_bs,
                            style=ft.ButtonStyle(color=ft.Colors.GREY_400),
                        ),
                    ],
                    alignment=ft.MainAxisAlignment.CENTER,
                    spacing=10,
                    wrap=True,
                )
            )
        else:
            # Just close button if no file
            content_controls.append(ft.Container(height=20))
            content_controls.append(
                ft.Row(
                    controls=[
                        ft.ElevatedButton(
                            "إغلاق",
                            on_click=close_bs,
                            bgcolor=ft.Colors.GREEN_700,
                            color=ft.Colors.WHITE,
                        ),
                    ],
                    alignment=ft.MainAxisAlignment.CENTER,
                )
            )
        
        content = ft.Column(
            controls=content_controls,
            horizontal_alignment=ft.CrossAxisAlignment.CENTER,
            spacing=10,
        )
        
        # Build header
        header_controls = [
            ft.Icon(ft.Icons.CHECK_CIRCLE, color=ft.Colors.GREEN_400, size=28),
            ft.Text(title, weight=ft.FontWeight.BOLD, size=20, color=ft.Colors.GREEN_300),
            ft.Container(expand=True),
            ft.IconButton(
                icon=ft.Icons.CLOSE,
                on_click=close_bs,
                icon_color=ft.Colors.GREY_400,
            ),
        ]
        
        bs = ft.BottomSheet(
            content=ft.Container(
                content=ft.Column(
                    controls=[
                        # Header
                        ft.Container(
                            content=ft.Row(
                                controls=header_controls,
                                alignment=ft.MainAxisAlignment.START,
                            ),
                            padding=ft.padding.only(left=20, right=20, top=20, bottom=10),
                        ),
                        ft.Divider(height=1, color=BottomSheetTheme.DIVIDER),
                        # Content
                        ft.Container(
                            content=content,
                            padding=ft.padding.all(20),
                        ),
                    ],
                    tight=True,
                ),
                bgcolor=BottomSheetTheme.BG,
            ),
            open=True,
            on_dismiss=close_bs,
        )
        
        page.overlay.append(bs)
        page.update()
        
        return bs
