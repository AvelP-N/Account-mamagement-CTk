import re
import subprocess
import tkinter as tk
from threading import Thread
import customtkinter


class PowerShellCommand:
    """A class for executing PowerShell commands"""

    @staticmethod
    def find_group_domain(ad_group):
        """Find a group in a domain"""

        command = f"""Get-ADGroup -Identity '{ad_group}'"""

        result_command = subprocess.run(['powershell.exe', '-Command', command], shell=False, capture_output=True,
                                        text=True, encoding="866", creationflags=subprocess.CREATE_NO_WINDOW)

        if result_command.returncode == 0:
            return "A group has been found in the domain"
        elif result_command.returncode == 1:
            return result_command.stderr.splitlines()[0]

    @staticmethod
    def add_domain_users(ad_group, ad_user):
        """Add users to a group with the PowerShell command"""

        command = f"""Add-ADGroupMember -Identity "{ad_group}" -Members "{ad_user}" -Confirm:$false"""

        result_command = subprocess.run(["powershell.exe", "-Command", command], shell=False, capture_output=True,
                                        text=True, encoding="866", creationflags=subprocess.CREATE_NO_WINDOW)

        if result_command.returncode == 0:
            return f"User < {ad_user} > have been added to the group\n{ad_group}"
        elif result_command.returncode == 1:
            return result_command.stderr.splitlines()[0]

    @staticmethod
    def remove_domain_users(ad_group, ad_user):
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

        self.group = None
        self.users = []

        self.title("Account Management  v1.1")
        self.resizable(False, False)
        # customtkinter.set_appearance_mode("dark")

        self.top_label = customtkinter.CTkLabel(self, text="Account and Group Management",
                                                font=("Times New Roman", 22, "bold"), text_color="gray")
        self.top_label.grid(row=0, column=0, columnspan=2, padx=10, pady=10)

        # The first frame to add users to the group
        self.main_frame = customtkinter.CTkFrame(self)
        self.main_frame.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")

        # CTkTextbox it is located on the right to display information about the actions performed
        self.textbox = customtkinter.CTkTextbox(self, activate_scrollbars=True, width=380, text_color="green")
        self.textbox.grid(row=1, column=1, padx=(0, 10), pady=10, sticky="nsew")

        # The frame where the group will be added
        self.frame_add_group = customtkinter.CTkFrame(self.main_frame)
        self.frame_add_group.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")

        self.entry_group = customtkinter.CTkEntry(self.frame_add_group, placeholder_text="Enter the group", width=250)
        self.entry_group.grid(row=0, column=0, padx=10, pady=10)
        self.entry_group.bind('<Return>', lambda event: self.add_group(self.entry_group.get()))
        self.entry_group.bind("<Button-3>", self.on_right_click_group)

        self.button_add_group = customtkinter.CTkButton(self.frame_add_group, text="Add group", text_color="black",
                                                        font=("Times New Roman", 14, "bold"),
                                                        command=lambda: self.add_group(self.entry_group.get()))
        self.button_add_group.grid(row=0, column=1, padx=10, pady=10)

        self.label_custom_text = customtkinter.CTkLabel(self.frame_add_group, text="")
        self.label_custom_text.grid(row=1, column=0, columnspan=2, padx=10, pady=10, sticky="w")

        self.button_delete_group = customtkinter.CTkButton(self.frame_add_group, text="Delete group",
                                                           text_color="black", font=("Times New Roman", 14, "bold"),
                                                           command=self.delete_group)
        self.button_delete_group.grid(row=1, column=1, padx=10, pady=10, sticky="e")

        # The frame where users will be added
        self.frame_add_user = customtkinter.CTkFrame(self.main_frame)
        self.frame_add_user.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")

        self.textbox_users = customtkinter.CTkTextbox(self.frame_add_user, activate_scrollbars=True, width=250,
                                                      height=80)
        self.textbox_users.grid(row=0, column=0, padx=10, pady=10, rowspan=2, sticky="nsew")
        self.textbox_users.bind("<Button-3>", self.on_right_click_users)

        self.button_add_users = customtkinter.CTkButton(self.frame_add_user, text="Add users", text_color="black",
                                                        font=("Times New Roman", 14, "bold"),
                                                        command=lambda:
                                                        self.add_users(self.textbox_users.get("1.0", "end-1c")))
        self.button_add_users.grid(row=0, column=1, padx=10, pady=10, sticky="e")

        self.button_delete_users = customtkinter.CTkButton(self.frame_add_user, text="Delete users", text_color="black",
                                                           font=("Times New Roman", 14, "bold"),
                                                           command=self.delete_users)
        self.button_delete_users.grid(row=1, column=1, padx=10, pady=10, sticky="e")

        # A frame for adding or removing users, a selected group
        self.frame_apply = customtkinter.CTkFrame(self.main_frame)
        self.frame_apply.grid(row=2, column=0, padx=10, pady=10, sticky="nsew")

        self.label_add_users = customtkinter.CTkLabel(self.frame_apply, text="Add users to a group",
                                                      text_color="grey", font=("Times New Roman", 16, "bold"),
                                                      width=250)
        self.label_add_users.grid(row=0, column=0, padx=10, pady=10, sticky="w")

        self.button_add_users = customtkinter.CTkButton(self.frame_apply, text="Add", text_color="black",
                                                        font=("Times New Roman", 14, "bold"),
                                                        command=self.add_users_to_a_group)
        self.button_add_users.grid(row=0, column=1, padx=10, pady=10)

        self.label_dell_users = customtkinter.CTkLabel(self.frame_apply, text="Remove users from a group",
                                                       text_color="grey", font=("Times New Roman", 16, "bold"),
                                                       width=250)
        self.label_dell_users.grid(row=1, column=0, padx=10, pady=10, sticky="w")

        self.button_dell_users = customtkinter.CTkButton(self.frame_apply, text="Delete", text_color="black",
                                                         font=("Times New Roman", 14, "bold"),
                                                         command=self.remove_users_from_a_group)
        self.button_dell_users.grid(row=1, column=1, padx=10, pady=10)

    def on_right_click_group(self, event):
        """Inserts the copied data into the input field for the group by right-clicking"""

        data = self.clipboard_get()
        self.entry_group.insert(tk.INSERT, data)

    def on_right_click_users(self, event):
        """Inserts the copied data into the user input field, right-click"""

        data = self.clipboard_get()
        self.textbox_users.insert(tk.INSERT, data)

    def add_group(self, group):
        """The button for adding a group to a variable"""

        if group:
            self.group = group.strip()
            self.entry_group.delete(0, customtkinter.END)
            self.label_custom_text.configure(text=group[:32], text_color="green", font=("Times New Roman", 16, "bold"))

            self.textbox.insert(customtkinter.END, f"***  The group has been added  ***\n{group}\n\n\n")
            self.textbox.see(customtkinter.END)

        else:
            self.textbox.insert(customtkinter.END, "Enter a group in the input field!\n\n")
            self.textbox.see(customtkinter.END)

    def delete_group(self):
        """A button to remove a group from a variable"""

        if self.group:
            self.label_custom_text.configure(text="")

            self.textbox.insert(customtkinter.END, f"***  The group has been deleted  ***\n{self.group}\n\n\n")
            self.textbox.see(customtkinter.END)
            self.group = None
        else:
            self.textbox.insert(customtkinter.END, "There is no added group!\n\n")
            self.textbox.see(customtkinter.END)

    def add_users(self, users):
        """A button to add users to the list"""

        self.textbox_users.delete("1.0", "end")

        if users:
            listing_users = set(re.findall(r"(?ai)([a-z]+(?:_[a-z]+|\.[a-z]+)?)(?:@[a-z]+\.[a-z]+)?", users))

            if listing_users:
                self.users.extend(listing_users)

                self.textbox.insert(customtkinter.END, "***  Users have been added to the list  ***\n")

                for user in self.users:
                    self.textbox.insert(customtkinter.END, f"{user}\n")

                self.textbox.insert(customtkinter.END, "\n\n")
                self.textbox.see(customtkinter.END)

            else:
                self.textbox.insert(customtkinter.END, "No users found!\n\n")
                self.textbox.see(customtkinter.END)

        else:
            self.textbox.insert(customtkinter.END, "No users found!\n\n")
            self.textbox.see(customtkinter.END)

    def delete_users(self):
        """A button to remove users from the list"""

        if self.users:
            self.users = []
            self.textbox.insert(customtkinter.END, "***  The list of users has been deleted  ***\n\n")
            self.textbox.see(customtkinter.END)
        else:
            self.textbox.insert(customtkinter.END, "The list of users is empty!\n\n")
            self.textbox.see(customtkinter.END)

    def add_users_to_a_group(self):
        """A button to add users to the selected group"""

        def add_users_in_thread(group, received_user):
            """Starts a thread to add a user to a group"""

            result_add_users = self.add_domain_users(group, received_user)

            self.textbox.insert(customtkinter.END, f"{result_add_users}\n\n")
            self.textbox.see(customtkinter.END)

        if self.group and self.users:
            find_group = self.find_group_domain(self.group)

            if find_group in "A group has been found in the domain":
                self.textbox.insert(customtkinter.END, f"***  Users will be added to the group < {self.group} >  "
                                                       f"***"f"\n\n")
                self.textbox.see(customtkinter.END)

                # Adding users to a selected group in multithreaded mode
                for user in self.users:
                    Thread(target=add_users_in_thread, args=(self.group, user)).start()

                self.textbox.insert(customtkinter.END, "\n")
                self.textbox.see(customtkinter.END)

            else:
                self.textbox.insert(customtkinter.END, f"{find_group}\n\n")
                self.textbox.see(customtkinter.END)

        else:
            self.textbox.insert(customtkinter.END, "Add a group and users!\n\n")
            self.textbox.see(customtkinter.END)

    def remove_users_from_a_group(self):
        """A button to remove users from the selected group"""

        def remove_users_in_thread(group, received_user):
            """Starts a thread to remove a user from a group"""

            result_remove_users = self.remove_domain_users(group, received_user)

            self.textbox.insert(customtkinter.END, f"{result_remove_users}\n\n")
            self.textbox.see(customtkinter.END)

        if self.group and self.users:
            find_group = self.find_group_domain(self.group)

            if find_group in "A group has been found in the domain":
                self.textbox.insert(customtkinter.END, f"***  Users will be removed from the group < {self.group} >"
                                                       f"  ***\n\n")
                self.textbox.see(customtkinter.END)

                # Deleting users to a selected group in multithreaded mode
                for user in self.users:
                    Thread(target=remove_users_in_thread, args=(self.group, user)).start()

                self.textbox.insert(customtkinter.END, "\n")
                self.textbox.see(customtkinter.END)

            else:
                self.textbox.insert(customtkinter.END, f"{find_group}\n\n")
                self.textbox.see(customtkinter.END)

        else:
            self.textbox.insert(customtkinter.END, "Add a group and users!\n\n")
            self.textbox.see(customtkinter.END)


if __name__ == '__main__':
    app = App()
    app.mainloop()
