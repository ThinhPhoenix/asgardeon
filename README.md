### ⚙️ 𝓐𝓼𝓰𝓪𝓻𝓭𝓮𝓸𝓷: _Windows Tweak Tool_

![Asgardeon](https://github.com/user-attachments/assets/20a2f6ed-5a48-4357-bfe0-c7d6ebcdda52)

**Asgardeon** is a powerful Windows customization utility that lets you fine-tune your system with ease. Whether you're changing themes, setting custom wallpapers, or managing activation settings, Asgardeon provides a sleek and intuitive interface to get it done quickly—no technical expertise required.

---

### 🚀 𝓕𝓮𝓪𝓽𝓾𝓻𝓮𝓼

- 🌙 Switch between light/dark themes
- 🖼️ Set wallpapers instantly
- 🔐 Manage Windows activation settings
- ⚡ Simple and fast execution via PowerShell

---

### 🛠️ 𝓗𝓸𝔀 𝓽𝓸 𝓾𝓼𝓮

1. Open **PowerShell as Administrator**
2. Run the following command to launch the full UI:

```powershell
irm https://l8.nu/125QH | iex
```

<details>
<summary><code>🔍 𝓞𝓻 𝓾𝓼𝓮 𝓽𝓱𝓮 𝓯𝓾𝓵𝓵 𝓤𝓡𝓛 𝓫𝓮𝓵𝓸𝔀</code></summary>

```powershell
irm https://raw.githubusercontent.com/ThinhPhoenix/asgardeon/refs/heads/main/main.ps1 | iex
```

</details>

### 📋 𝓓𝓲𝓻𝓮𝓬𝓽 𝓕𝓾𝓷𝓬𝓽𝓲𝓸𝓷 𝓤𝓼𝓪𝓰𝓮

You can also invoke specific functions directly:

#### Set Dark Theme:
```powershell
irm https://l8.nu/125QH | iex; Set-AsgardeonTheme -Theme Dark
```

#### Set Light Theme:
```powershell
irm https://l8.nu/125QH | iex; Set-AsgardeonTheme -Theme Light
```

#### Set Desktop Wallpaper:
```powershell
irm https://l8.nu/125QH | iex; Set-AsgardeonWallpaper -ImagePath "C:\path\to\wallpaper.jpg"
```

#### Run Windows/Office Activation:
```powershell
irm https://l8.nu/125QH | iex; Start-WindowsActivation
```

#### Hide Windows Activation Watermark:
```powershell
irm https://l8.nu/125QH | iex; Enable-HideActivation
```
