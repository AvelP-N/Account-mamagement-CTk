import re
import subprocess
import tkinter as tk
from threading import Thread
import customtkinter


class PowerShellCommand:
    """A class for executing PowerShell commands"""

    @staticmethod
    def find_group_domain(ad_group: str) -> str:
        """Find a group in a domain"""

        command = f"Get-ADGroup -Identity '{ad_group}'"

        result_command = subprocess.run(['powershell.exe', '-Command', command], shell=False, capture_output=True,
                                        text=True, encoding="866", creationflags=subprocess.CREATE_NO_WINDOW)

        if result_command.returncode == 0:
            return "A group has been found in the domain"
        elif result_command.returncode == 1:
            return result_command.stderr.splitlines()[0]

    @staticmethod
    def find_user_domain(ad_user: str) -> str:
        """Search for a user in the domain"""

        command = f"Get-ADUser -Identity '{ad_user}'"

        result_command = subprocess.run(['powershell.exe', '-Command', command], shell=False, capture_output=True,
                                        text=True, encoding="866", creationflags=subprocess.CREATE_NO_WINDOW)

        if result_command.returncode == 0:
            return "The user has been found in the domain"
        elif result_command.returncode == 1:
            return result_command.stderr.splitlines()[0]

    @staticmethod
    def add_domain_user(ad_group: str, ad_user: str) -> str:
        """Add user to a group with the PowerShell command"""

        command = f"""Add-ADGroupMember -Identity "{ad_group}" -Members "{ad_user}" -Confirm:$false"""

        result_command = subprocess.run(["powershell.exe", "-Command", command], shell=False, capture_output=True,
                                        text=True, encoding="866", creationflags=subprocess.CREATE_NO_WINDOW)

        if result_command.returncode == 0:
            return f"User < {ad_user} > have been added to the group:\n{ad_group}"
        elif result_command.returncode == 1:
            return result_command.stderr.splitlines()[0]

    @staticmethod
    def remove_domain_user(ad_group: str, ad_user: str) -> str:
        """Deleting user from a group with the PowerShell command"""

        command = f"""Remove-ADGroupMember -Identity "{ad_group}" -Members "{ad_user}" -Confirm:$false"""

        result_command = subprocess.run(["powershell.exe", "-Command", command], shell=False, capture_output=True,
                                        text=True, encoding="866", creationflags=subprocess.CREATE_NO_WINDOW)

        if result_command.returncode == 0:
            return f"The user < {ad_user} > has been removed from the group\n{ad_group}"
        elif result_command.returncode == 1:
            return result_command.stderr.splitlines()[0]


class App(customtkinter.CTk, PowerShellCommand):
    """A class for rendering the program"""

    def __init__(self):
        super().__init__()

        self.group = None
        self.user = None
        self.list_users = []
        self.list_groups = []

        # *** Setup main interface ***
        self.title("Account Management  v1.3")
        self.resizable(False, False)
        # customtkinter.set_appearance_mode("dark")

        self.top_label = customtkinter.CTkLabel(self, text="Account and Group Management",
                                                font=("Times New Roman", 22, "bold"), text_color="gray")
        self.top_label.grid(row=0, column=0, columnspan=2, padx=10, pady=10)

        # Create tabview
        self.tabview = customtkinter.CTkTabview(self)
        self.tabview.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")
        self.tabview.add("Adding users to a group")
        self.tabview.add("Adding groups to a user")

        # CTkTextbox it is located on the right to display information about the actions performed
        self.textbox = customtkinter.CTkTextbox(self, activate_scrollbars=True, width=400, text_color="green")
        self.textbox.grid(row=1, column=1, padx=(0, 10), pady=(27, 8), sticky="nsew")

        # *** The main frame in the "Adding users to a group" tab ***
        self.main_frame_tab1 = customtkinter.CTkFrame(self.tabview.tab("Adding users to a group"))
        self.main_frame_tab1.grid(row=0, column=0)

        # A frame for adding a group
        self.frame_add_group = customtkinter.CTkFrame(self.main_frame_tab1)
        self.frame_add_group.grid(row=0, column=0, pady=(0, 10), sticky="nsew")

        self.entry_group = customtkinter.CTkEntry(self.frame_add_group, width=250,
                                                  placeholder_text="Enter the group")
        self.entry_group.grid(row=0, column=0, padx=10, pady=10)
        self.entry_group.bind('<Return>', lambda event: self.add_group(self.entry_group.get()))
        self.entry_group.bind("<Button-3>", self.on_right_click_group)

        self.button_add_group = customtkinter.CTkButton(self.frame_add_group, text="Add group",
                                                        text_color="black", font=("Times New Roman", 16, "bold"),
                                                        command=lambda: self.add_group(self.entry_group.get()))
        self.button_add_group.grid(row=0, column=1, padx=10, pady=10)

        self.label_text_group = customtkinter.CTkLabel(self.frame_add_group, text="")
        self.label_text_group.grid(row=1, column=0, columnspan=2, padx=10, pady=10, sticky="w")

        self.button_delete_group = customtkinter.CTkButton(self.frame_add_group, text="Delete data", text_color="black",
                                                           font=("Times New Roman", 16, "bold"),
                                                           command=self.delete_group)
        self.button_delete_group.grid(row=1, column=1, padx=10, pady=10, sticky="e")

        # A frame for adding users
        self.frame_add_users = customtkinter.CTkFrame(self.main_frame_tab1)
        self.frame_add_users.grid(row=1, column=0, pady=(0, 10), sticky="nsew")

        self.textbox_users = customtkinter.CTkTextbox(self.frame_add_users, activate_scrollbars=True,
                                                      width=250, height=80)
        self.textbox_users.grid(row=0, column=0, padx=10, pady=10, rowspan=2, sticky="nsew")
        self.textbox_users.bind("<Button-3>", self.on_right_click_users)

        self.button_add_users = customtkinter.CTkButton(self.frame_add_users, text="Add users",
                                                        text_color="black", font=("Times New Roman", 16, "bold"),
                                                        command=lambda:
                                                        self.add_list_users(self.textbox_users.get("1.0", "end-1c")))
        self.button_add_users.grid(row=0, column=1, padx=10, pady=10, sticky="e")

        self.button_delete_users = customtkinter.CTkButton(self.frame_add_users, text="Delete data", text_color="black",
                                                           font=("Times New Roman", 16, "bold"),
                                                           command=self.delete_list_users)
        self.button_delete_users.grid(row=1, column=1, padx=10, pady=10, sticky="e")

        # A frame for adding or removing users, a selected group
        self.frame_apply_users = customtkinter.CTkFrame(self.main_frame_tab1)
        self.frame_apply_users.grid(row=2, column=0, sticky="nsew")

        self.label_add_users = customtkinter.CTkLabel(self.frame_apply_users, text="Add a list of users to a group",
                                                      text_color="grey", font=("Times New Roman", 16, "bold"),
                                                      width=250)
        self.label_add_users.grid(row=0, column=0, padx=10, pady=10, sticky="w")

        self.button_add_users = customtkinter.CTkButton(self.frame_apply_users, text="Add", text_color="black",
                                                        font=("Times New Roman", 16, "bold"),
                                                        command=self.add_users_to_a_group)
        self.button_add_users.grid(row=0, column=1, padx=10, pady=10)

        self.label_dell_users = customtkinter.CTkLabel(self.frame_apply_users,
                                                       text="Delete a list of users from a group",
                                                       text_color="grey", font=("Times New Roman", 16, "bold"),
                                                       width=250)
        self.label_dell_users.grid(row=1, column=0, padx=10, pady=10, sticky="w")

        self.button_dell_users = customtkinter.CTkButton(self.frame_apply_users, text="Delete", text_color="black",
                                                         font=("Times New Roman", 16, "bold"),
                                                         command=self.remove_users_from_a_group)
        self.button_dell_users.grid(row=1, column=1, padx=10, pady=10)

        # *** The main frame in the "Adding groups to a user" tab ***
        self.main_frame_tab2 = customtkinter.CTkFrame(self.tabview.tab("Adding groups to a user"))
        self.main_frame_tab2.grid(row=0, column=0)

        # A frame for adding a user
        self.frame_add_user = customtkinter.CTkFrame(self.main_frame_tab2)
        self.frame_add_user.grid(row=0, column=0, pady=(0, 10), sticky="nsew")

        self.entry_user = customtkinter.CTkEntry(self.frame_add_user, width=250,
                                                 placeholder_text="Enter the user")
        self.entry_user.grid(row=0, column=0, padx=10, pady=10)
        self.entry_user.bind('<Return>', lambda event: self.add_user(self.entry_user.get()))
        self.entry_user.bind("<Button-3>", self.on_right_click_user)

        self.button_add_user = customtkinter.CTkButton(self.frame_add_user, text="Add user",
                                                       text_color="black", font=("Times New Roman", 16, "bold"),
                                                       command=lambda: self.add_user(self.entry_user.get()))
        self.button_add_user.grid(row=0, column=1, padx=10, pady=10)

        self.label_text_user = customtkinter.CTkLabel(self.frame_add_user, text="")
        self.label_text_user.grid(row=1, column=0, columnspan=2, padx=10, pady=10, sticky="w")

        self.button_delete_user = customtkinter.CTkButton(self.frame_add_user, text="Delete data", text_color="black",
                                                          font=("Times New Roman", 16, "bold"),
                                                          command=self.delete_user)
        self.button_delete_user.grid(row=1, column=1, padx=10, pady=10, sticky="e")

        # A frame for adding groups
        self.frame_add_groups = customtkinter.CTkFrame(self.main_frame_tab2)
        self.frame_add_groups.grid(row=1, column=0, pady=(0, 10), sticky="nsew")

        self.textbox_groups = customtkinter.CTkTextbox(self.frame_add_groups, activate_scrollbars=True,
                                                       width=250, height=80)
        self.textbox_groups.grid(row=0, column=0, padx=10, pady=10, rowspan=2, sticky="nsew")
        self.textbox_groups.bind("<Button-3>", self.on_right_click_groups)

        self.button_add_groups = customtkinter.CTkButton(self.frame_add_groups, text="Add groups",
                                                         text_color="black", font=("Times New Roman", 16, "bold"),
                                                         command=lambda:
                                                         self.add_list_groups(self.textbox_groups.get("1.0", "end-1c")))
        self.button_add_groups.grid(row=0, column=1, padx=10, pady=10, sticky="e")

        self.button_delete_groups = customtkinter.CTkButton(self.frame_add_groups, text="Delete data",
                                                            text_color="black", font=("Times New Roman", 16, "bold"),
                                                            command=self.delete_list_groups)
        self.button_delete_groups.grid(row=1, column=1, padx=10, pady=10, sticky="e")

        # A frame for adding or removing a user from selected groups
        self.frame_apply_groups = customtkinter.CTkFrame(self.main_frame_tab2)
        self.frame_apply_groups.grid(row=2, column=0, sticky="nsew")

        self.label_add_groups = customtkinter.CTkLabel(self.frame_apply_groups, text="Add a list of groups to the user",
                                                       text_color="grey", font=("Times New Roman", 16, "bold"),
                                                       width=250)
        self.label_add_groups.grid(row=0, column=0, padx=10, pady=10, sticky="w")

        self.button_add_groups = customtkinter.CTkButton(self.frame_apply_groups, text="Add", text_color="black",
                                                         font=("Times New Roman", 16, "bold"),
                                                         command=self.add_groups_to_a_user)
        self.button_add_groups.grid(row=0, column=1, padx=10, pady=10)

        self.label_dell_groups = customtkinter.CTkLabel(self.frame_apply_groups, text_color="grey",
                                                        text="Remove a user from the list of groups",
                                                        font=("Times New Roman", 14, "bold"), width=250)
        self.label_dell_groups.grid(row=1, column=0, padx=10, pady=10, sticky="w")

        self.button_dell_groups = customtkinter.CTkButton(self.frame_apply_groups, text="Delete", text_color="black",
                                                          font=("Times New Roman", 16, "bold"),
                                                          command=self.remove_a_user_from_groups)
        self.button_dell_groups.grid(row=1, column=1, padx=10, pady=10)

    def on_right_click_group(self, event: tk.Event) -> None:
        """Inserts the copied data into the input field to add a group by right-clicking"""

        data = self.clipboard_get()
        self.entry_group.insert(tk.INSERT, data)

    def on_right_click_users(self, event: tk.Event) -> None:
        """Inserts the copied data into the input field to add users to the list, right-click"""

        data = self.clipboard_get()
        self.textbox_users.insert(tk.INSERT, data)

    def add_group(self, data: str) -> None:
        """The button for adding a group to a variable"""

        if data:
            group = re.findall(r"[^@]{2,}", data)

            if group:
                self.group = str(group[0]).strip()
                self.entry_group.delete(0, customtkinter.END)
                self.label_text_group.configure(text=self.group[:32], text_color="green",
                                                font=("Times New Roman", 16, "bold"))
                self.write_text(f"***  The group has been added  ***\n{self.group}\n\n\n")

            else:
                self.write_text("The group was not found!\n\n")

        else:
            self.write_text("Enter a group in the input field!\n\n")

    def delete_group(self) -> None:
        """A button to remove a group from a variable"""

        if self.group:
            self.label_text_group.configure(text="")
            self.write_text(f"***  The data has been deleted  ***\n{self.group}\n\n\n")
            self.group = None

        else:
            self.write_text("There is no data to delete!\n\n")

    def add_list_users(self, data: str) -> None:
        """A button to add users to the list"""

        self.textbox_users.delete("1.0", "end")

        if data:
            self.list_users = list(self.find_groups_or_users(data))

            if self.list_users:
                self.write_text("***  Users have been added to the list  ***\n")

                for data in self.list_users:
                    self.write_text(f"{data}\n")

                self.write_text("\n\n")

            else:
                self.write_text("No users found!\n\n")

        else:
            self.write_text("Enter data in the input field!\n\n")

    @staticmethod
    def find_groups_or_users(data_input_field: str) -> set:
        """Users or groups are searched from the data in the text field"""

        input_data = data_input_field
        found_data = []

        # Search for ntfs groups
        pattern_find_ntfs_groups = re.compile(r"(?i)(?:croc-|msk-|perm-|QVSRV_).+?(?:_ro|_rw)\b")
        list_groups_ntfs = pattern_find_ntfs_groups.findall(input_data)

        if list_groups_ntfs:
            found_data.extend(list_groups_ntfs)

            pattern_ntfs = r"|".join(list_groups_ntfs)
            input_data = re.sub(pattern_ntfs, "", input_data)

        # Search for rds groups
        rds_groups = ("some groups")

        list_groups_rds = [group for group in rds_groups if group in input_data]

        if list_groups_rds:
            found_data.extend(list_groups_rds)

            pattern_rds = r"|".join(list_groups_rds)
            input_data = re.sub(pattern_rds, "", input_data)

        input_data = re.sub(r"KDL|OC|IT|SSL|KIPS|Taxcom", "", input_data)

        # Search for other groups or users
        pattern_find_rest_groups_or_users = re.compile(r"(?ai)(\w{2,}(?:[.-]\w+)*)(?:@[a-z]+\.[a-z]+)?")
        list_rest_data = pattern_find_rest_groups_or_users.findall(input_data)
        found_data.extend(list_rest_data)

        return set(found_data)

    def delete_list_users(self) -> None:
        """A button to remove users from the list"""

        if self.list_users:
            self.list_users.clear()
            self.write_text("***  The list of users has been deleted  ***\n\n")

        else:
            self.write_text("There is no data to delete!\n\n")

    def add_users_to_a_group(self) -> None:
        """A button to add users to the selected group"""

        def add_user_in_thread(received_user: str) -> None:
            """Starts a thread to add a user to a group"""

            result_add_user = self.add_domain_user(self.group, received_user)

            self.write_text(f"{result_add_user}\n\n")

        if self.group and self.list_users:
            find_group = self.find_group_domain(self.group)

            if find_group in "A group has been found in the domain":
                self.write_text("\n***  Adding a list of users to a group  ***\n\n")

                # Adding users to a selected group in multithreaded mode
                for user in self.list_users:
                    Thread(target=add_user_in_thread, args=(user,)).start()

            else:
                self.write_text(f"{find_group}\n\n")

        else:
            self.write_text("Add a group and a list of users!\n\n")

    def remove_users_from_a_group(self) -> None:
        """A button to remove users from the selected group"""

        def remove_user_in_thread(received_user: str) -> None:
            """Starts a thread to remove a user from a group"""

            result_remove_user = self.remove_domain_user(self.group, received_user)

            self.write_text(f"{result_remove_user}\n\n")

        if self.group and self.list_users:
            find_group = self.find_group_domain(self.group)

            if find_group in "A group has been found in the domain":
                self.write_text("\n***  Deleting a list of users from a group  ***\n\n")

                # Deleting users to a selected group in multithreaded mode
                for user in self.list_users:
                    Thread(target=remove_user_in_thread, args=(user,)).start()

            else:
                self.write_text(f"{find_group}\n\n")

        else:
            self.write_text("Add a group and a list of users!\n\n")

    def on_right_click_user(self, event: tk.Event) -> None:
        """Inserts the copied data into the input field to add a user by right-clicking"""

        data = self.clipboard_get()
        self.entry_user.insert(tk.INSERT, data)

    def on_right_click_groups(self, event: tk.Event) -> None:
        """Inserts the copied data into the input field to add groups to the list, right-click"""

        data = self.clipboard_get()
        self.textbox_groups.insert(tk.INSERT, data)

    def add_user(self, data: str) -> None:
        """The button for adding a user to a variable"""

        if data:
            user = re.findall(r"[^@ А-ЯЁа-яё]{2,}", data)

            if user:
                self.user = str(user[0])
                self.entry_user.delete(0, customtkinter.END)
                self.label_text_user.configure(text=self.user[:32], text_color="green",
                                               font=("Times New Roman", 16, "bold"))
                self.write_text(f"***  The user has been added  ***\n{self.user}\n\n\n")

            else:
                self.entry_user.delete(0, customtkinter.END)
                self.write_text("The user was not found!\n\n")

        else:
            self.write_text("Enter a user in the input field!\n\n")

    def delete_user(self) -> None:
        """A button to remove a user from a variable"""

        if self.user:
            self.label_text_user.configure(text="")
            self.write_text(f"***  The data has been deleted  ***\n{self.user}\n\n\n")
            self.user = None

        else:
            self.write_text("There is no data to delete!\n\n")

    def add_list_groups(self, data: str) -> None:
        """A button to add groups to the list"""

        self.textbox_groups.delete("1.0", "end")

        if data:
            self.list_groups = list(self.find_groups_or_users(data))

            if self.list_groups:
                self.write_text("***  Groups have been added to the list  ***\n")

                for data in self.list_groups:
                    self.write_text(f"{data}\n")

                self.write_text("\n\n")

            else:
                self.write_text("No groups found!\n\n")

        else:
            self.write_text("Enter data in the input field!\n\n")

    def delete_list_groups(self) -> None:
        """A button to remove groups from the list"""

        if self.list_groups:
            self.list_groups.clear()
            self.write_text("***  The list of groups has been deleted  ***\n\n")

        else:
            self.write_text("There is no data to delete!\n\n")

    def add_groups_to_a_user(self) -> None:
        """A button to add a user to the list of groups"""

        def add_group_in_thread(received_group: str) -> None:
            """starts a thread to add a group to the user"""

            result_add_group = self.add_domain_user(received_group, self.user)

            self.write_text(f"{result_add_group}\n\n")

        if self.user and self.list_groups:
            find_user = self.find_user_domain(self.user)

            if find_user in "The user has been found in the domain":
                self.write_text("\n***  Adding a user to the list of groups  ***\n\n")

                # Adding groups to a user in multithreaded mode
                for group in self.list_groups:
                    Thread(target=add_group_in_thread, args=(group,)).start()

            else:
                self.write_text(f"{find_user}\n\n")

        else:
            self.write_text("Add a user and a list of groups!\n\n")

    def remove_a_user_from_groups(self) -> None:
        """A button to remove a user from the list of groups"""

        def remove_group_in_thread(received_group: str) -> None:
            """Starts a thread to delete a group from a user"""

            result_remove_user = self.remove_domain_user(received_group, self.user)

            self.write_text(f"{result_remove_user}\n\n")

        if self.user and self.list_groups:
            find_user = self.find_user_domain(self.user)

            if find_user in "The user has been found in the domain":
                self.write_text(f"\n***  Removing a user from the list of groups  ***\n\n")

                # Deleting groups from a user in multithreaded mode
                for group in self.list_groups:
                    Thread(target=remove_group_in_thread, args=(group,)).start()

            else:
                self.write_text(f"{find_user}\n\n")

        else:
            self.write_text("Add a user and a list of groups!\n\n")

    def write_text(self, text: str) -> None:
        self.textbox.insert(customtkinter.END, text)
        self.textbox.see(customtkinter.END)


if __name__ == '__main__':
    app = App()
    app.mainloop()


