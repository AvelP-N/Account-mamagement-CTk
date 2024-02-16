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
    def add_domain_users(ad_group: str, ad_user: str) -> str:
        """Add users to a group with the PowerShell command"""

        command = f"""Add-ADGroupMember -Identity "{ad_group}" -Members "{ad_user}" -Confirm:$false"""

        result_command = subprocess.run(["powershell.exe", "-Command", command], shell=False, capture_output=True,
                                        text=True, encoding="866", creationflags=subprocess.CREATE_NO_WINDOW)

        if result_command.returncode == 0:
            return f"User < {ad_user} > have been added to the group:\n{ad_group}"
        elif result_command.returncode == 1:
            return result_command.stderr.splitlines()[0]

    @staticmethod
    def remove_domain_users(ad_group: str, ad_user: str) -> str:
        """Deleting users from a group with the PowerShell command"""

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

        self.group_or_user = None
        self.list_users_or_groups = []

        # Setup main interface
        self.title("Account Management  v1.2")
        self.resizable(False, False)
        # customtkinter.set_appearance_mode("dark")

        self.top_label = customtkinter.CTkLabel(self, text="Account and Group Management",
                                                font=("Times New Roman", 22, "bold"), text_color="gray")
        self.top_label.grid(row=0, column=0, columnspan=2, padx=10, pady=10)

        # Main frame
        self.main_frame = customtkinter.CTkFrame(self)
        self.main_frame.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")

        # CTkTextbox it is located on the right to display information about the actions performed
        self.textbox = customtkinter.CTkTextbox(self, activate_scrollbars=True, width=400, text_color="green")
        self.textbox.grid(row=1, column=1, padx=(0, 10), pady=10, sticky="nsew")

        # A frame for adding a group or user
        self.frame_add_group_or_user = customtkinter.CTkFrame(self.main_frame)
        self.frame_add_group_or_user.grid(row=0, column=0, padx=10, pady=(10, 5), sticky="nsew")

        self.entry_group_or_user = customtkinter.CTkEntry(self.frame_add_group_or_user, width=250,
                                                          placeholder_text="Enter a group or user")
        self.entry_group_or_user.grid(row=0, column=0, padx=10, pady=10)
        self.entry_group_or_user.bind('<Return>', lambda event: self.add_group_or_user(self.entry_group_or_user.get()))
        self.entry_group_or_user.bind("<Button-3>", self.on_right_click_group_or_user)

        self.button_add_group_or_user = customtkinter.CTkButton(self.frame_add_group_or_user, text="Add group or user",
                                                                text_color="black", font=("Times New Roman", 14, "bold"),
                                                                command=lambda:
                                                                self.add_group_or_user(self.entry_group_or_user.get()))
        self.button_add_group_or_user.grid(row=0, column=1, padx=10, pady=10)

        self.label_custom_text = customtkinter.CTkLabel(self.frame_add_group_or_user, text="")
        self.label_custom_text.grid(row=1, column=0, columnspan=2, padx=10, pady=10, sticky="w")

        self.button_delete_group_or_user = customtkinter.CTkButton(self.frame_add_group_or_user, text="Delete data",
                                                                   text_color="black",
                                                                   font=("Times New Roman", 14, "bold"),
                                                                   command=self.delete_group_or_user)
        self.button_delete_group_or_user.grid(row=1, column=1, padx=10, pady=10, sticky="e")

        # A frame for adding users or groups to the list
        self.frame_add_users_or_groups = customtkinter.CTkFrame(self.main_frame)
        self.frame_add_users_or_groups.grid(row=1, column=0, padx=10, pady=5, sticky="nsew")

        self.textbox_users_or_groups = customtkinter.CTkTextbox(self.frame_add_users_or_groups, activate_scrollbars=True,
                                                                width=250, height=80)
        self.textbox_users_or_groups.grid(row=0, column=0, padx=10, pady=10, rowspan=2, sticky="nsew")
        self.textbox_users_or_groups.bind("<Button-3>", self.on_right_click_users_or_groups)

        self.button_add_users_or_groups = customtkinter.CTkButton(self.frame_add_users_or_groups,
                                                                  text="Add users or groups", text_color="black",
                                                                  font=("Times New Roman", 14, "bold"),
                                                                  command=lambda: self.add_users_or_groups(self.textbox_users_or_groups.get("1.0", "end-1c")))
        self.button_add_users_or_groups.grid(row=0, column=1, padx=10, pady=10, sticky="e")

        self.button_delete_users_or_groups = customtkinter.CTkButton(self.frame_add_users_or_groups, text="Delete data",
                                                                     text_color="black",
                                                                     font=("Times New Roman", 14, "bold"),
                                                                     command=self.delete_users_or_groups)
        self.button_delete_users_or_groups.grid(row=1, column=1, padx=10, pady=10, sticky="e")

        # A frame for adding or removing users, a selected group
        self.frame_apply_users = customtkinter.CTkFrame(self.main_frame)
        self.frame_apply_users.grid(row=2, column=0, padx=10, pady=5, sticky="nsew")

        self.label_add_users = customtkinter.CTkLabel(self.frame_apply_users, text="Add a list of users to a group",
                                                      text_color="grey", font=("Times New Roman", 16, "bold"),
                                                      width=250)
        self.label_add_users.grid(row=0, column=0, padx=10, pady=10, sticky="w")

        self.button_add_users = customtkinter.CTkButton(self.frame_apply_users, text="Add", text_color="black",
                                                        font=("Times New Roman", 14, "bold"),
                                                        command=self.add_users_to_a_group)
        self.button_add_users.grid(row=0, column=1, padx=10, pady=10)

        self.label_dell_users = customtkinter.CTkLabel(self.frame_apply_users,
                                                       text="Delete a list of users from a group",
                                                       text_color="grey", font=("Times New Roman", 16, "bold"),
                                                       width=250)
        self.label_dell_users.grid(row=1, column=0, padx=10, pady=10, sticky="w")

        self.button_dell_users = customtkinter.CTkButton(self.frame_apply_users, text="Delete", text_color="black",
                                                         font=("Times New Roman", 14, "bold"),
                                                         command=self.remove_users_from_a_group)
        self.button_dell_users.grid(row=1, column=1, padx=10, pady=10)

        # A frame for adding or removing a user from selected groups
        self.frame_apply_groups = customtkinter.CTkFrame(self.main_frame)
        self.frame_apply_groups.grid(row=3, column=0, padx=10, pady=(5, 10), sticky="nsew")

        self.label_add_users = customtkinter.CTkLabel(self.frame_apply_groups, text="Add a user to the list of groups",
                                                      text_color="grey", font=("Times New Roman", 16, "bold"),
                                                      width=250)
        self.label_add_users.grid(row=0, column=0, padx=10, pady=10, sticky="w")

        self.button_add_users = customtkinter.CTkButton(self.frame_apply_groups, text="Add", text_color="black",
                                                        font=("Times New Roman", 14, "bold"),
                                                        command=self.add_groups_to_a_user)
        self.button_add_users.grid(row=0, column=1, padx=10, pady=10)

        self.label_dell_users = customtkinter.CTkLabel(self.frame_apply_groups, text_color="grey",
                                                       text="Remove a user from the list of groups",
                                                       font=("Times New Roman", 16, "bold"),
                                                       width=250)
        self.label_dell_users.grid(row=1, column=0, padx=10, pady=10, sticky="w")

        self.button_dell_users = customtkinter.CTkButton(self.frame_apply_groups, text="Delete", text_color="black",
                                                         font=("Times New Roman", 14, "bold"),
                                                         command=self.remove_a_user_from_groups)
        self.button_dell_users.grid(row=1, column=1, padx=10, pady=10)

    def on_right_click_group_or_user(self, event: tk.Event) -> None:
        """Inserts the copied data into the input field to add a group or user by right-clicking"""

        data = self.clipboard_get()
        self.entry_group_or_user.insert(tk.INSERT, data)

    def on_right_click_users_or_groups(self, event: tk.Event) -> None:
        """Inserts the copied data into the input field to add users or groups to the list, right-click"""

        data = self.clipboard_get()
        self.textbox_users_or_groups.insert(tk.INSERT, data)

    def add_group_or_user(self, data: str) -> None:
        """The button for adding a group or user to a variable"""

        if data:
            group_or_user = re.findall(r"[^@]+", data)

            if group_or_user:
                self.group_or_user = str(group_or_user[0]).strip()
                self.entry_group_or_user.delete(0, customtkinter.END)
                self.label_custom_text.configure(text=self.group_or_user[:32],
                                                 text_color="green", font=("Times New Roman", 16, "bold"))

                self.write_text(f"***  Data added  ***\n{self.group_or_user}\n\n\n")

        else:
            self.write_text("Enter a group or user in the input field!\n\n")

    def delete_group_or_user(self) -> None:
        """A button to remove a group or user from a variable"""

        if self.group_or_user:
            self.label_custom_text.configure(text="")

            self.write_text(f"***  The data has been deleted  ***\n{self.group_or_user}\n\n\n")

            self.group_or_user = None

        else:
            self.write_text("There is no added group or user!\n\n")

    def add_users_or_groups(self, data: str) -> None:
        """A button to add users or groups to the list"""

        self.textbox_users_or_groups.delete("1.0", "end")

        if data:
            self.list_users_or_groups = list(self.find_groups_or_users(data))

            if self.list_users_or_groups:
                self.write_text("***  Data have been added to the list  ***\n")

                for data in self.list_users_or_groups:
                    self.write_text(f"{data}\n")

                self.write_text("\n\n")

            else:
                self.write_text("No users or group found!\n\n")

        else:
            self.write_text("Enter users or groups in the input field!\n\n")

    @staticmethod
    def find_groups_or_users(data_input_field: str) -> set:
        """Users or groups are searched from the data in the text field"""

        input_data = data_input_field
        found_data = []

        # Search for ntfs groups
        pattern_find_ntfs_groups = re.compile(r"(?i)(?:croc-|msk-|perm-|QVSRV_).+?(?:_ro|_rw)\b")
        list_groups_ntfs = pattern_find_ntfs_groups.findall(input_data)
        found_data.extend(list_groups_ntfs)

        pattern_ntfs = r"|".join(list_groups_ntfs)
        input_data = re.sub(pattern_ntfs, "", input_data)

        # Search for rds groups
        rds_groups = ("some groups")

        list_groups_rds = [group for group in rds_groups if group in input_data]
        found_data.extend(list_groups_rds)

        pattern_rds = r"|".join(list_groups_rds) + "|KDL|OC"
        input_data = re.sub(pattern_rds, "", input_data)

        # Search for other groups or users
        pattern_find_rest_groups_or_users = re.compile(r"(?ai)(\w+(?:[.-]\w+)*)(?:@[a-z]+\.[a-z]+)?")
        list_rest_data = pattern_find_rest_groups_or_users.findall(input_data)
        found_data.extend(list_rest_data)

        return set(found_data)

    def delete_users_or_groups(self) -> None:
        """A button to remove users or groups from the list"""

        if self.list_users_or_groups:
            self.list_users_or_groups = []
            self.write_text("***  The list of users or groups has been deleted  ***\n\n")

        else:
            self.write_text("The list of users or groups is empty!\n\n")

    def add_users_to_a_group(self) -> None:
        """A button to add users to the selected group"""

        def add_users_in_thread(received_user: str) -> None:
            """Starts a thread to add a user to a group"""

            result_add_user = self.add_domain_users(self.group_or_user, received_user)

            self.write_text(f"{result_add_user}\n\n")

        if self.group_or_user and self.list_users_or_groups:
            find_group = self.find_group_domain(self.group_or_user)

            if find_group in "A group has been found in the domain":
                self.write_text(f"***  Users will be added to the group < {self.group_or_user} >  ***\n\n")

                # Adding users to a selected group in multithreaded mode
                for user in self.list_users_or_groups:
                    Thread(target=add_users_in_thread, args=(user,)).start()

                self.write_text("\n")

            else:
                self.write_text(f"{find_group}\n\n")

        else:
            self.write_text("Add a group and a list of users!\n\n")

    def remove_users_from_a_group(self) -> None:
        """A button to remove users from the selected group"""

        def remove_users_in_thread(received_user: str) -> None:
            """Starts a thread to remove a user from a group"""

            result_remove_user = self.remove_domain_users(self.group_or_user, received_user)

            self.write_text(f"{result_remove_user}\n\n")

        if self.group_or_user and self.list_users_or_groups:
            find_group = self.find_group_domain(self.group_or_user)

            if find_group in "A group has been found in the domain":
                self.write_text(f"***  Users will be removed from the group < {self.group_or_user} >  ***\n\n")

                # Deleting users to a selected group in multithreaded mode
                for user in self.list_users_or_groups:
                    Thread(target=remove_users_in_thread, args=(user,)).start()

                self.write_text("\n")

            else:
                self.write_text(f"{find_group}\n\n")

        else:
            self.write_text("Add a group and a list of users!\n\n")

    def add_groups_to_a_user(self) -> None:
        """A button to add a user to the list of groups"""

        def add_groups_in_thread(received_group: str) -> None:
            """starts a thread to add a group to the user"""

            result_add_group = self.add_domain_users(received_group, self.group_or_user)

            self.write_text(f"{result_add_group}\n\n")

        if self.group_or_user and self.list_users_or_groups:
            find_user = self.find_user_domain(self.group_or_user)

            if find_user in "The user has been found in the domain":
                self.write_text(f"***  The groups will be added to the user < {self.group_or_user} >  ***\n\n")

                # Adding groups to a user in multithreaded mode
                for group in self.list_users_or_groups:
                    Thread(target=add_groups_in_thread, args=(group,)).start()

                self.write_text("\n")

            else:
                self.write_text(f"{find_user}\n\n")

        else:
            self.write_text("Add a user and a list of groups!\n\n")

    def remove_a_user_from_groups(self) -> None:
        """A button to remove a user from the list of groups"""

        def remove_groups_in_thread(received_group: str) -> None:
            """Starts a thread to delete a group from a user"""

            result_remove_user = self.remove_domain_users(received_group, self.group_or_user)

            self.write_text(f"{result_remove_user}\n\n")

        if self.group_or_user and self.list_users_or_groups:
            find_user = self.find_user_domain(self.group_or_user)

            if find_user in "The user has been found in the domain":
                self.write_text(f"***  The user < {self.group_or_user} > will be removed from the groups  ***\n\n")

                # Deleting groups from a user in multithreaded mode
                for group in self.list_users_or_groups:
                    Thread(target=remove_groups_in_thread, args=(group,)).start()

                self.write_text("\n")

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
