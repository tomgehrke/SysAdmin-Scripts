$NetworkAdapterProperties = @(
    "IPv4 Checksum Offload",
    "IPv4 TSO Offload",
    "Large Send offload v2 (ipv4)",
    "Large Send offload v2 (ipv6)",
    "TCP checksum offload (ipv4)",
    "TCP checksum offload (ipv6)",
    "UDP checksum offload (ipv4)",
    "UDP checksum offload (ipv6)",
    "Offload ip options",
    "Offload tcp options",
    "Offload tagged traffic"
)

foreach ($Property in $NetworkAdapterProperties) {
    Get-NetAdapterAdvancedProperty -DisplayName $Property -ErrorAction SilentlyContinue
}
